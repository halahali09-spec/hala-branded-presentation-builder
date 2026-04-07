// ============================================================
// BRANDED PRESENTATION BUILDER — app.js
// Upload branded .pptx → extract REAL backgrounds/logo/colors/fonts
// Paste content → generate presentation with actual template backgrounds
// ============================================================

// ---- State ----
let brands = JSON.parse(localStorage.getItem('pptx-brands') || '[]');
let activeBrand = null;
let parsedSlides = [];

// ---- DOM refs ----
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

const stepEls = $$('.step');
const panels = { 1: $('#step1'), 2: $('#step2'), 3: $('#step3') };

const dropZone = $('#drop-zone');
const pptxUpload = $('#pptx-upload');
const extractStatus = $('#extract-status');
const brandPreview = $('#brand-preview');
const brandNameInput = $('#brand-name');
const savedBrandsList = $('#saved-brands-list');
const saveBrandBtn = $('#save-brand-btn');
const useBrandBtn = $('#use-brand-btn');
const manualLogo = $('#manual-logo');
const bgPreviewRow = $('#bg-preview-row');

const contentInput = $('#content-input');
const parsedPreviewEl = $('#parsed-preview');
const generateBtn = $('#generate-btn');
const activeBrandBadge = $('#active-brand-badge');
const formatGuideToggle = $('#format-guide-toggle');
const formatGuideBody = $('#format-guide-body');

const previewContainer = $('#preview-container');
const downloadBtn = $('#download-btn');

// ============================================================
// NAVIGATION
// ============================================================
let currentStep = 1;

function goToStep(n) {
    currentStep = n;
    Object.values(panels).forEach(p => p.classList.remove('active'));
    panels[n].classList.add('active');
    stepEls.forEach(el => {
        const s = parseInt(el.dataset.step);
        el.classList.remove('active', 'done');
        if (s === n) el.classList.add('active');
        else if (s < n) el.classList.add('done');
    });
    if (n === 3) renderPreview();
}

$('#back-to-step1').addEventListener('click', () => goToStep(1));
$('#back-to-step2').addEventListener('click', () => goToStep(2));

// ============================================================
// STEP 1 — EXTRACT BRANDING FROM .PPTX
// ============================================================

dropZone.addEventListener('click', () => pptxUpload.click());
dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file && file.name.endsWith('.pptx')) handlePptxUpload(file);
});
pptxUpload.addEventListener('change', (e) => {
    if (e.target.files[0]) handlePptxUpload(e.target.files[0]);
});

manualLogo.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const dataUrl = await fileToDataURL(file);
    const logoBox = $('#extracted-logo-box');
    logoBox.innerHTML = `<img src="${dataUrl}" alt="Logo">`;
    logoBox.dataset.logo = dataUrl;
});

async function handlePptxUpload(file) {
    extractStatus.classList.remove('hidden');
    brandPreview.classList.add('hidden');
    try {
        const brand = await extractBrandFromPptx(file);
        showBrandPreview(brand);
    } catch (err) {
        console.error('Extraction error:', err);
        alert('Could not read branding from this file. Error: ' + err.message);
    }
    extractStatus.classList.add('hidden');
}

// ============================================================
// PPTX EXTRACTION ENGINE
// ============================================================

async function extractBrandFromPptx(file) {
    const zip = await JSZip.loadAsync(await file.arrayBuffer());

    const brand = {
        name: file.name.replace('.pptx', '').replace(/[_-]/g, ' ').trim(),
        primaryColor: '#2E4DFF',
        secondaryColor: '#0F161A',
        accentColor: '#FF410C',
        textLight: '#FFFFFF',
        textDark: '#0F161A',
        headingFont: 'Arial',
        bodyFont: 'Arial',
        logo: null,
        backgrounds: {
            title: null,   // dark dramatic bg for title/cover slides
            dark: null,    // dark bg for content slides
            light: null,   // light bg for content slides
        },
    };

    // 1. Find the right theme (check all themes, pick the one with non-default colors)
    const themeColors = await extractBestThemeColors(zip);
    if (themeColors.accent1) brand.primaryColor = themeColors.accent1;
    if (themeColors.dk1) brand.secondaryColor = themeColors.dk1;
    if (themeColors.accent6 || themeColors.accent2) brand.accentColor = themeColors.accent6 || themeColors.accent2;
    if (themeColors.lt1) brand.textLight = themeColors.lt1;
    if (themeColors.dk1) brand.textDark = themeColors.dk1;
    if (themeColors.fonts) {
        if (themeColors.fonts.major) brand.headingFont = themeColors.fonts.major;
        if (themeColors.fonts.minor) brand.bodyFont = themeColors.fonts.minor;
    }

    // 2. Extract background images from slide layouts
    const backgrounds = await extractBackgrounds(zip);
    brand.backgrounds = backgrounds;

    // 3. Extract logo
    const logo = await extractLogo(zip);
    if (logo) brand.logo = logo;

    return brand;
}

async function extractBestThemeColors(zip) {
    // Check all themes, prefer the one with non-default Office colors
    const defaultAccent1 = '4472C4'; // Default Office blue
    let bestColors = {};
    let bestFonts = {};

    for (let i = 1; i <= 5; i++) {
        const themeFile = zip.file(`ppt/theme/theme${i}.xml`);
        if (!themeFile) continue;

        const xml = await themeFile.async('string');
        const colors = parseThemeColors(xml);
        const fonts = parseThemeFonts(xml);

        // If this theme has non-default accent1, prefer it
        if (colors.accent1 && colors.accent1.replace('#', '') !== defaultAccent1) {
            bestColors = colors;
            bestFonts = fonts;
            break;
        }

        // Store first theme as fallback
        if (i === 1) {
            bestColors = colors;
            bestFonts = fonts;
        }
    }

    bestColors.fonts = bestFonts;
    return bestColors;
}

function parseThemeColors(xmlString) {
    const colors = {};
    const names = ['dk1', 'dk2', 'lt1', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];
    for (const name of names) {
        const srgb = xmlString.match(new RegExp(`<a:${name}>\\s*<a:srgbClr val="([A-Fa-f0-9]{6})"`, 'i'));
        if (srgb) { colors[name] = '#' + srgb[1]; continue; }
        const sys = xmlString.match(new RegExp(`<a:${name}>\\s*<a:sysClr[^>]*lastClr="([A-Fa-f0-9]{6})"`, 'i'));
        if (sys) colors[name] = '#' + sys[1];
    }
    return colors;
}

function parseThemeFonts(xmlString) {
    const fonts = {};
    const major = xmlString.match(/<a:majorFont>[\s\S]*?<a:latin typeface="([^"]+)"/i);
    if (major) fonts.major = major[1];
    const minor = xmlString.match(/<a:minorFont>[\s\S]*?<a:latin typeface="([^"]+)"/i);
    if (minor) fonts.minor = minor[1];
    return fonts;
}

// ============================================================
// BACKGROUND EXTRACTION
// ============================================================

async function extractBackgrounds(zip) {
    const backgrounds = { title: null, dark: null, light: null };
    const layoutImages = [];
    const seenFilenames = new Set();

    // Helper to add an image candidate
    const addImage = async (filename, source) => {
        if (seenFilenames.has(filename)) return;
        seenFilenames.add(filename);
        if (!filename.match(/\.(png|jpg|jpeg)$/i)) return;

        const mediaFile = zip.file('ppt/media/' + filename);
        if (!mediaFile) return;

        const data = await mediaFile.async('arraybuffer');
        // Backgrounds are large images
        if (data.byteLength < 50000) return;

        const rawDataUrl = arrayBufferToDataURL(data, getMime(filename));
        // Compress to keep file size manageable for PptxGenJS
        const compressed = await compressImage(rawDataUrl, 1600, 0.82);

        layoutImages.push({
            source, filename,
            size: data.byteLength,
            dataUrl: compressed,
        });
        console.log(`Found bg: ${filename} (${(data.byteLength/1024).toFixed(0)}KB → compressed)`);
    };

    // Check slide layouts for background images
    for (let i = 1; i <= 30; i++) {
        const relsFile = zip.file(`ppt/slideLayouts/_rels/slideLayout${i}.xml.rels`);
        if (!relsFile) continue;
        const relsXml = await relsFile.async('string');
        const matches = [...relsXml.matchAll(/Target="\.\.\/media\/([^"]+)"/g)];
        for (const m of matches) await addImage(m[1], `layout${i}`);
    }

    // Check first slide for title background
    const slide1Rels = zip.file('ppt/slides/_rels/slide1.xml.rels');
    if (slide1Rels) {
        const relsXml = await slide1Rels.async('string');
        const matches = [...relsXml.matchAll(/Target="\.\.\/media\/([^"]+)"/g)];
        for (const m of matches) await addImage(m[1], 'slide1');
    }

    // Check slide masters too
    for (let m = 1; m <= 2; m++) {
        const relsFile = zip.file(`ppt/slideMasters/_rels/slideMaster${m}.xml.rels`);
        if (!relsFile) continue;
        const relsXml = await relsFile.async('string');
        const matches = [...relsXml.matchAll(/Target="\.\.\/media\/([^"]+)"/g)];
        for (const match of matches) await addImage(match[1], `master${m}`);
    }

    console.log(`Total backgrounds found: ${layoutImages.length}`);
    if (layoutImages.length === 0) return backgrounds;

    return await classifyBackgrounds(layoutImages);
}

// Compress image to reduce size
function compressImage(dataUrl, maxWidth, quality) {
    return new Promise((resolve) => {
        const img = new Image();
        img.onload = () => {
            const canvas = document.createElement('canvas');
            const scale = Math.min(1, maxWidth / img.width);
            canvas.width = Math.round(img.width * scale);
            canvas.height = Math.round(img.height * scale);
            const ctx = canvas.getContext('2d');
            ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
            // Use JPEG for backgrounds (smaller than PNG)
            resolve(canvas.toDataURL('image/jpeg', quality));
        };
        img.onerror = () => resolve(dataUrl);
        img.src = dataUrl;
    });
}

async function classifyBackgrounds(images) {
    const backgrounds = { title: null, dark: null, light: null };
    const analyzed = [];

    for (const img of images) {
        const brightness = await getImageBrightness(img.dataUrl);
        analyzed.push({ ...img, brightness });
    }

    // Sort by brightness
    analyzed.sort((a, b) => a.brightness - b.brightness);

    // Deduplicate by size (same size = likely same image)
    const unique = [];
    const seenSizes = new Set();
    for (const a of analyzed) {
        if (!seenSizes.has(a.size)) {
            unique.push(a);
            seenSizes.add(a.size);
        }
    }

    if (unique.length >= 3) {
        // Darkest = title, next dark = dark content, lightest = light content
        backgrounds.title = unique[0].dataUrl;
        backgrounds.dark = unique[1].dataUrl;
        backgrounds.light = unique[unique.length - 1].dataUrl;
    } else if (unique.length === 2) {
        backgrounds.title = unique[0].dataUrl;
        backgrounds.dark = unique[0].dataUrl;
        backgrounds.light = unique[1].dataUrl;
    } else if (unique.length === 1) {
        // Single background
        if (unique[0].brightness < 128) {
            backgrounds.title = unique[0].dataUrl;
            backgrounds.dark = unique[0].dataUrl;
        } else {
            backgrounds.light = unique[0].dataUrl;
        }
    }

    return backgrounds;
}

function getImageBrightness(dataUrl) {
    return new Promise((resolve) => {
        const img = new Image();
        img.onload = () => {
            const canvas = document.createElement('canvas');
            const size = 50; // Sample at small size for speed
            canvas.width = size;
            canvas.height = size;
            const ctx = canvas.getContext('2d');
            ctx.drawImage(img, 0, 0, size, size);
            const data = ctx.getImageData(0, 0, size, size).data;
            let total = 0;
            for (let i = 0; i < data.length; i += 4) {
                total += (data[i] + data[i + 1] + data[i + 2]) / 3;
            }
            resolve(total / (size * size));
        };
        img.onerror = () => resolve(128);
        img.src = dataUrl;
    });
}

// ============================================================
// LOGO EXTRACTION
// ============================================================

async function extractLogo(zip) {
    // Strategy: find images in slide master that are NOT full-size backgrounds
    // Also check the white-on-dark logo pattern (small wide images)
    const candidates = [];

    // Check slide masters
    for (let m = 1; m <= 2; m++) {
        const relsFile = zip.file(`ppt/slideMasters/_rels/slideMaster${m}.xml.rels`);
        if (!relsFile) continue;
        const relsXml = await relsFile.async('string');
        const matches = [...relsXml.matchAll(/Target="\.\.\/media\/([^"]+)"/g)];
        for (const match of matches) {
            const filename = match[1];
            if (!filename.match(/\.(png|jpg|jpeg)$/i)) continue;
            candidates.push(filename);
        }
    }

    // Check slide layouts for small images (logos tend to be embedded in layout backgrounds)
    // Also scan all media for the logo pattern: wide aspect ratio, small-medium file
    const allMedia = [];
    zip.folder('ppt/media').forEach((path, file) => {
        if (path.match(/\.(png|jpg|jpeg)$/i)) allMedia.push(path);
    });

    // Analyze candidates first, then fall back to scanning all media
    const searchList = candidates.length > 0 ? candidates : allMedia.slice(0, 20);

    let bestLogo = null;
    let bestScore = -1;

    for (const filename of searchList) {
        const mediaFile = zip.file('ppt/media/' + filename);
        if (!mediaFile) continue;

        const data = await mediaFile.async('arraybuffer');
        // Skip very large files (backgrounds) and very tiny files (dots/icons)
        if (data.byteLength > 500000 || data.byteLength < 500) continue;

        const dataUrl = arrayBufferToDataURL(data, getMime(filename));

        // Check dimensions
        const dims = await getImageDimensions(dataUrl);
        if (!dims) continue;

        // Logo heuristic: wide aspect ratio (width > height * 1.5), reasonable size
        const aspect = dims.width / dims.height;
        let score = 0;
        if (aspect > 2) score += 3;       // Very wide = likely text logo
        else if (aspect > 1.5) score += 2;
        if (dims.width > 200 && dims.width < 3000) score += 1;
        if (data.byteLength > 2000 && data.byteLength < 200000) score += 1;

        if (score > bestScore) {
            bestScore = score;
            bestLogo = dataUrl;
        }
    }

    return bestLogo;
}

function getImageDimensions(dataUrl) {
    return new Promise((resolve) => {
        const img = new Image();
        img.onload = () => resolve({ width: img.width, height: img.height });
        img.onerror = () => resolve(null);
        img.src = dataUrl;
    });
}

// ============================================================
// BRAND UI
// ============================================================

function showBrandPreview(brand) {
    $('#ex-primary').value = brand.primaryColor;
    $('#ex-secondary').value = brand.secondaryColor;
    $('#ex-accent').value = brand.accentColor;
    $('#ex-text').value = brand.textLight;
    $('#ex-textdark').value = brand.textDark;
    $('#ex-heading-font').value = brand.headingFont;
    $('#ex-body-font').value = brand.bodyFont;
    brandNameInput.value = brand.name;

    // Logo
    const logoBox = $('#extracted-logo-box');
    if (brand.logo) {
        logoBox.innerHTML = `<img src="${brand.logo}" alt="Logo">`;
        logoBox.dataset.logo = brand.logo;
    } else {
        logoBox.innerHTML = '<span class="no-logo-text">No logo found</span>';
        logoBox.dataset.logo = '';
    }

    // Background previews
    bgPreviewRow.innerHTML = '';
    const bgLabels = { title: 'Title / Cover', dark: 'Dark Content', light: 'Light Content' };
    for (const [key, label] of Object.entries(bgLabels)) {
        if (brand.backgrounds[key]) {
            bgPreviewRow.innerHTML += `
                <div class="bg-preview-thumb" data-bg="${key}">
                    <img src="${brand.backgrounds[key]}" alt="${label}">
                    <div class="bg-label">${label}</div>
                </div>`;
        }
    }
    if (!bgPreviewRow.innerHTML) {
        bgPreviewRow.innerHTML = '<p class="no-brands">No background images found. Solid colors will be used.</p>';
    }

    // Store backgrounds on a temp object for save/use
    brandPreview.dataset.backgrounds = JSON.stringify(brand.backgrounds);
    brandPreview.classList.remove('hidden');
}

function getBrandFromPreview() {
    const name = brandNameInput.value.trim();
    if (!name) { alert('Please enter a brand name.'); return null; }
    const logoBox = $('#extracted-logo-box');
    let backgrounds = { title: null, dark: null, light: null };
    try { backgrounds = JSON.parse(brandPreview.dataset.backgrounds || '{}'); } catch(e) {}

    return {
        name,
        primaryColor: $('#ex-primary').value,
        secondaryColor: $('#ex-secondary').value,
        accentColor: $('#ex-accent').value,
        textLight: $('#ex-text').value,
        textDark: $('#ex-textdark').value,
        headingFont: $('#ex-heading-font').value || 'Arial',
        bodyFont: $('#ex-body-font').value || 'Arial',
        logo: logoBox.dataset.logo || null,
        backgrounds,
    };
}

saveBrandBtn.addEventListener('click', () => {
    const b = getBrandFromPreview();
    if (!b) return;
    const existingIdx = brands.findIndex(x => x.name.toLowerCase() === b.name.toLowerCase());
    if (existingIdx >= 0) brands[existingIdx] = b;
    else brands.push(b);
    persistBrands();
    renderBrands();
});

useBrandBtn.addEventListener('click', () => {
    const b = getBrandFromPreview();
    if (!b) return;
    activeBrand = b;
    activeBrandBadge.textContent = activeBrand.name;
    goToStep(2);
});

function persistBrands() {
    localStorage.setItem('pptx-brands', JSON.stringify(brands));
}

function renderBrands() {
    if (brands.length === 0) {
        savedBrandsList.innerHTML = '<p class="no-brands">No saved brands yet. Upload a .pptx to get started!</p>';
        return;
    }
    savedBrandsList.innerHTML = brands.map((b, i) => `
        <div class="brand-card" data-index="${i}">
            <button class="brand-card-delete" data-index="${i}" title="Delete">&times;</button>
            <div class="brand-card-name">${escapeHTML(b.name)}</div>
            <div class="brand-card-colors">
                <div class="brand-card-swatch" style="background:${b.primaryColor}"></div>
                <div class="brand-card-swatch" style="background:${b.secondaryColor}"></div>
                <div class="brand-card-swatch" style="background:${b.accentColor}"></div>
            </div>
        </div>
    `).join('');

    savedBrandsList.querySelectorAll('.brand-card').forEach(card => {
        card.addEventListener('click', (e) => {
            if (e.target.classList.contains('brand-card-delete')) return;
            showBrandPreview(brands[parseInt(card.dataset.index)]);
        });
    });
    savedBrandsList.querySelectorAll('.brand-card-delete').forEach(btn => {
        btn.addEventListener('click', (e) => {
            e.stopPropagation();
            brands.splice(parseInt(btn.dataset.index), 1);
            persistBrands();
            renderBrands();
        });
    });
}

renderBrands();

// ============================================================
// STEP 2 — CONTENT PARSING
// ============================================================

formatGuideToggle.addEventListener('click', () => {
    formatGuideBody.classList.toggle('hidden');
    formatGuideToggle.classList.toggle('open');
});

contentInput.addEventListener('input', () => {
    parsedSlides = parseContent(contentInput.value);
    renderParsedPreview();
    generateBtn.disabled = parsedSlides.length === 0;
});

function parseContent(text) {
    if (!text.trim()) return [];

    // Split by --- or double blank lines
    const rawSections = text.split(/\n\s*---\s*\n|\n\s*---\s*$|^\s*---\s*\n/);
    const slides = [];

    for (const section of rawSections) {
        const trimmed = section.trim();
        if (!trimmed) continue;

        // Further split by double newlines within a section (each becomes a slide)
        const subSections = trimmed.split(/\n\s*\n/);

        for (const sub of subSections) {
            const t = sub.trim();
            if (!t) continue;

            const lines = t.split('\n');
            const firstLine = lines[0].trim();

            // Title slide: # heading (single hash)
            if (firstLine.startsWith('# ') && !firstLine.startsWith('## ')) {
                const title = firstLine.replace(/^# /, '');
                const rest = lines.slice(1).map(l => l.trim()).filter(l => l).join('\n');
                slides.push({
                    type: isThankYou(title) ? 'thank-you' : 'title',
                    bgType: 'title',
                    data: { title, subtitle: rest }
                });
                continue;
            }

            // Content with heading: ## heading
            if (firstLine.startsWith('## ')) {
                const heading = firstLine.replace(/^## /, '');
                const bodyLines = lines.slice(1).map(l => l.trim()).filter(l => l);
                const bodyText = bodyLines.join('\n');

                // Two columns
                if (bodyText.match(/LEFT:/i) && bodyText.match(/RIGHT:/i)) {
                    const leftMatch = bodyText.match(/LEFT:\s*([\s\S]*?)(?=RIGHT:)/i);
                    const rightMatch = bodyText.match(/RIGHT:\s*([\s\S]*)/i);
                    slides.push({
                        type: 'two-column', bgType: 'light',
                        data: {
                            heading,
                            left: leftMatch ? leftMatch[1].trim() : '',
                            right: rightMatch ? rightMatch[1].trim() : '',
                        }
                    });
                    continue;
                }

                // Bullets
                const allBullets = bodyLines.length > 0 && bodyLines.every(l => l.startsWith('- ') || l.startsWith('* '));
                if (allBullets) {
                    slides.push({
                        type: 'bullets', bgType: 'light',
                        data: { heading, bullets: bodyLines.map(l => l.replace(/^[-*] /, '')).join('\n') }
                    });
                    continue;
                }

                // Regular content
                slides.push({
                    type: 'content', bgType: 'light',
                    data: { heading, body: bodyLines.join('\n') }
                });
                continue;
            }

            // Plain text without any heading marker
            // Check if it has bullets
            const allBullets = lines.every(l => l.trim().startsWith('- ') || l.trim().startsWith('* '));
            if (allBullets && lines.length > 1) {
                slides.push({
                    type: 'bullets', bgType: 'light',
                    data: { heading: '', bullets: lines.map(l => l.trim().replace(/^[-*] /, '')).join('\n') }
                });
            } else {
                slides.push({
                    type: 'content', bgType: 'light',
                    data: { heading: '', body: t }
                });
            }
        }
    }

    return slides;
}

function isThankYou(text) {
    const lower = text.toLowerCase();
    return lower.includes('thank') || lower.includes('questions') || lower.includes('q&a') || lower.includes('the end');
}

function renderParsedPreview() {
    if (parsedSlides.length === 0) {
        parsedPreviewEl.innerHTML = '';
        return;
    }
    const labels = { 'title': 'Title', 'thank-you': 'Thank You', 'content': 'Content', 'bullets': 'Bullets', 'two-column': 'Two Columns' };
    parsedPreviewEl.innerHTML =
        `<p style="font-size:13px;font-weight:600;color:#475569;margin-bottom:8px;">${parsedSlides.length} slide${parsedSlides.length > 1 ? 's' : ''} detected:</p>` +
        parsedSlides.map((s, i) => `
            <div class="parsed-slide-card">
                <div class="parsed-slide-num">${i + 1}</div>
                <div>
                    <div class="parsed-slide-type">${labels[s.type] || s.type}</div>
                    <div class="parsed-slide-title">${escapeHTML(s.data.title || s.data.heading || s.data.body?.slice(0, 60) || 'Untitled')}</div>
                </div>
            </div>`).join('');
}

generateBtn.addEventListener('click', () => {
    parsedSlides = parseContent(contentInput.value);
    if (parsedSlides.length > 0) goToStep(3);
});

// ============================================================
// STEP 3 — PREVIEW & DOWNLOAD
// ============================================================

function renderPreview() {
    const b = activeBrand;
    previewContainer.innerHTML = parsedSlides.map((s, i) => {
        const isDark = s.bgType === 'title' || s.bgType === 'dark';
        const textColor = isDark ? b.textLight : b.textDark;
        const headingColor = isDark ? b.textLight : b.primaryColor;

        // Background: use extracted image if available, else solid color
        let bgStyle = '';
        const bgImg = b.backgrounds[s.bgType];
        if (bgImg) {
            bgStyle = `background-image:url('${bgImg}');background-size:cover;background-position:center;`;
        } else {
            bgStyle = isDark ? `background:${b.secondaryColor};` : `background:#F5F6F7;`;
        }

        let inner = '';
        if (s.type === 'title' || s.type === 'thank-you') {
            inner = `
                <div style="flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;">
                    <div style="font-size:28px;font-weight:700;color:${headingColor};font-family:'${b.headingFont}',sans-serif;">${escapeHTML(s.data.title || '')}</div>
                    <div style="font-size:16px;color:${textColor};margin-top:12px;font-family:'${b.bodyFont}',sans-serif;opacity:0.85;">${escapeHTML(s.data.subtitle || '')}</div>
                </div>`;
        } else if (s.type === 'content') {
            inner = `
                <div style="font-size:22px;font-weight:700;color:${headingColor};font-family:'${b.headingFont}',sans-serif;margin-bottom:16px;">${escapeHTML(s.data.heading || '')}</div>
                <div style="font-size:14px;color:${textColor};font-family:'${b.bodyFont}',sans-serif;white-space:pre-wrap;flex:1;opacity:0.9;">${escapeHTML(s.data.body || '')}</div>`;
        } else if (s.type === 'bullets') {
            const bullets = (s.data.bullets || '').split('\n').filter(x => x.trim()).map(x => `<li style="margin-bottom:6px;">${escapeHTML(x)}</li>`).join('');
            inner = `
                <div style="font-size:22px;font-weight:700;color:${headingColor};font-family:'${b.headingFont}',sans-serif;margin-bottom:16px;">${escapeHTML(s.data.heading || '')}</div>
                <ul style="font-size:14px;color:${textColor};font-family:'${b.bodyFont}',sans-serif;padding-left:24px;flex:1;">${bullets}</ul>`;
        } else if (s.type === 'two-column') {
            inner = `
                <div style="font-size:22px;font-weight:700;color:${headingColor};font-family:'${b.headingFont}',sans-serif;margin-bottom:16px;">${escapeHTML(s.data.heading || '')}</div>
                <div style="display:flex;gap:24px;flex:1;">
                    <div style="flex:1;font-size:13px;color:${textColor};font-family:'${b.bodyFont}',sans-serif;white-space:pre-wrap;">${escapeHTML(s.data.left || '')}</div>
                    <div style="flex:1;font-size:13px;color:${textColor};font-family:'${b.bodyFont}',sans-serif;white-space:pre-wrap;">${escapeHTML(s.data.right || '')}</div>
                </div>`;
        }

        const logoTag = b.logo ? `<img class="preview-slide-logo" src="${b.logo}">` : '';

        return `
        <div class="preview-slide" style="${bgStyle}">
            <span class="preview-slide-number">Slide ${i + 1}</span>
            ${inner}
            ${logoTag}
        </div>`;
    }).join('');
}

// -- DOWNLOAD PPTX --
downloadBtn.addEventListener('click', () => {
    const b = activeBrand;
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE';
    pptx.author = 'Branded Presentation Builder';

    const hex = (c) => (c || '#000000').replace('#', '');

    parsedSlides.forEach((s) => {
        const slide = pptx.addSlide();
        const isDark = s.bgType === 'title' || s.bgType === 'dark';
        const textColor = isDark ? hex(b.textLight) : hex(b.textDark);
        const headingColor = isDark ? hex(b.textLight) : hex(b.primaryColor);

        // Background: image if available, else solid color
        const bgImg = b.backgrounds[s.bgType];
        if (bgImg) {
            // PptxGenJS expects data without the "data:" URI prefix
            const bgData = bgImg.startsWith('data:') ? bgImg.substring(5) : bgImg;
            slide.background = { data: bgData };
        } else {
            slide.background = { color: isDark ? hex(b.secondaryColor) : 'F5F6F7' };
        }

        // Logo on every slide (bottom-left like Drata template)
        if (b.logo) {
            slide.addImage({
                data: b.logo,
                x: 0.4, y: 6.8,
                w: 1.2, h: 0.4,
                sizing: { type: 'contain', w: 1.2, h: 0.4 },
            });
        }

        if (s.type === 'title' || s.type === 'thank-you') {
            slide.addText(s.data.title || '', {
                x: 0.8, y: 2.0, w: 11.5, h: 1.5,
                fontSize: 40, fontFace: b.headingFont,
                color: headingColor,
                bold: true, align: 'center', valign: 'middle',
            });
            if (s.data.subtitle) {
                slide.addText(s.data.subtitle, {
                    x: 0.8, y: 3.6, w: 11.5, h: 1.0,
                    fontSize: 18, fontFace: b.bodyFont,
                    color: textColor,
                    align: 'center', valign: 'top',
                });
            }
        } else if (s.type === 'content') {
            if (s.data.heading) {
                slide.addText(s.data.heading, {
                    x: 0.8, y: 0.5, w: 11.5, h: 0.9,
                    fontSize: 30, fontFace: b.headingFont,
                    color: headingColor, bold: true,
                });
            }
            slide.addText(s.data.body || '', {
                x: 0.8, y: s.data.heading ? 1.6 : 0.8, w: 11.5, h: s.data.heading ? 5.0 : 5.8,
                fontSize: 16, fontFace: b.bodyFont,
                color: textColor, valign: 'top',
            });
        } else if (s.type === 'bullets') {
            if (s.data.heading) {
                slide.addText(s.data.heading, {
                    x: 0.8, y: 0.5, w: 11.5, h: 0.9,
                    fontSize: 30, fontFace: b.headingFont,
                    color: headingColor, bold: true,
                });
            }
            const items = (s.data.bullets || '').split('\n').filter(x => x.trim()).map(text => ({
                text,
                options: { bullet: { type: 'bullet' }, fontSize: 16, fontFace: b.bodyFont, color: textColor },
            }));
            if (items.length > 0) {
                slide.addText(items, {
                    x: 0.8, y: s.data.heading ? 1.6 : 0.8, w: 11.5, h: s.data.heading ? 5.0 : 5.8,
                    valign: 'top',
                });
            }
        } else if (s.type === 'two-column') {
            if (s.data.heading) {
                slide.addText(s.data.heading, {
                    x: 0.8, y: 0.5, w: 11.5, h: 0.9,
                    fontSize: 30, fontFace: b.headingFont,
                    color: headingColor, bold: true,
                });
            }
            const startY = s.data.heading ? 1.6 : 0.8;
            const h = s.data.heading ? 5.0 : 5.8;
            slide.addText(s.data.left || '', {
                x: 0.8, y: startY, w: 5.4, h: h,
                fontSize: 15, fontFace: b.bodyFont, color: textColor, valign: 'top',
            });
            slide.addText(s.data.right || '', {
                x: 6.9, y: startY, w: 5.4, h: h,
                fontSize: 15, fontFace: b.bodyFont, color: textColor, valign: 'top',
            });
        }
    });

    const fileName = (b.name + ' Presentation').replace(/[^a-zA-Z0-9 _-]/g, '');
    pptx.writeFile({ fileName });
});

// ============================================================
// UTILITIES
// ============================================================

function escapeHTML(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}

function fileToDataURL(file) {
    return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(e.target.result);
        reader.readAsDataURL(file);
    });
}

function arrayBufferToDataURL(buffer, mime) {
    const bytes = new Uint8Array(buffer);
    let binary = '';
    for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
    return 'data:' + mime + ';base64,' + btoa(binary);
}

function getMime(filename) {
    const ext = filename.split('.').pop().toLowerCase();
    return { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif' }[ext] || 'image/png';
}
