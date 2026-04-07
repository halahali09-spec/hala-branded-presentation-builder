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
    const logoResult = await extractLogo(zip);
    if (logoResult) {
        brand.logo = logoResult.dataUrl;
        brand.logoIsLight = logoResult.isLight;
    }

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
    // NEW STRATEGY: walk through the template's actual slides IN ORDER and grab
    // the background image of each one. We then cycle through this ordered list
    // when generating, so the output visually mirrors the template's flow.
    const orderedBackgrounds = []; // [{ dataUrl, brightness, source }, ...]
    const filenameCache = new Map(); // filename -> { dataUrl, brightness } (cached compress)

    const loadAndCompress = async (filename) => {
        if (filenameCache.has(filename)) return filenameCache.get(filename);
        if (!filename.match(/\.(png|jpg|jpeg)$/i)) return null;
        const mediaFile = zip.file('ppt/media/' + filename);
        if (!mediaFile) return null;
        const data = await mediaFile.async('arraybuffer');
        // Backgrounds are large images (skip tiny icons)
        if (data.byteLength < 30000) return null;
        const rawDataUrl = arrayBufferToDataURL(data, getMime(filename));
        const compressed = await compressImage(rawDataUrl, 1600, 0.82);
        const brightness = await getImageBrightness(compressed);
        const entry = { dataUrl: compressed, brightness, filename };
        filenameCache.set(filename, entry);
        return entry;
    };

    // Find background image filename referenced by a layout's rels (largest image)
    const layoutBgCache = new Map();
    const getLayoutBackground = async (layoutNum) => {
        if (layoutBgCache.has(layoutNum)) return layoutBgCache.get(layoutNum);
        const relsFile = zip.file(`ppt/slideLayouts/_rels/slideLayout${layoutNum}.xml.rels`);
        if (!relsFile) { layoutBgCache.set(layoutNum, null); return null; }
        const relsXml = await relsFile.async('string');
        const matches = [...relsXml.matchAll(/Target="\.\.\/media\/([^"]+)"/g)];
        // Pick the LARGEST image referenced (the background, not small icons)
        let best = null;
        for (const m of matches) {
            const entry = await loadAndCompress(m[1]);
            if (entry) { best = entry; break; } // first valid large image is usually the background
        }
        layoutBgCache.set(layoutNum, best);
        return best;
    };

    // Walk through ALL slides in the template in order
    let slideIdx = 1;
    while (true) {
        const slideRels = zip.file(`ppt/slides/_rels/slide${slideIdx}.xml.rels`);
        if (!slideRels) break;
        const relsXml = await slideRels.async('string');

        // First try: slide has its own background image embedded
        const imageMatches = [...relsXml.matchAll(/Target="\.\.\/media\/([^"]+)"/g)];
        let chosen = null;
        for (const m of imageMatches) {
            const entry = await loadAndCompress(m[1]);
            if (entry) { chosen = entry; break; }
        }

        // Fallback: use the slide's layout background
        if (!chosen) {
            const layoutMatch = relsXml.match(/slideLayouts\/slideLayout(\d+)\.xml/);
            if (layoutMatch) {
                chosen = await getLayoutBackground(parseInt(layoutMatch[1], 10));
            }
        }

        if (chosen) {
            orderedBackgrounds.push({ ...chosen, slideIdx });
            console.log(`Slide ${slideIdx}: ${chosen.filename} (brightness=${chosen.brightness.toFixed(0)})`);
        }
        slideIdx++;
        if (slideIdx > 100) break; // safety
    }

    // Deduplicate consecutive identical backgrounds (template often repeats the same layout)
    // but KEEP order — just remove exact duplicates that are next to each other
    const deduped = [];
    let lastFilename = null;
    for (const bg of orderedBackgrounds) {
        if (bg.filename !== lastFilename) {
            deduped.push(bg);
            lastFilename = bg.filename;
        }
    }

    console.log(`Extracted ${deduped.length} unique background(s) from ${orderedBackgrounds.length} template slides`);

    // Build BOTH the ordered sequence AND the legacy {title, dark, light} categorization
    // for backward compat with the preview UI
    const sequence = deduped.map(b => ({ dataUrl: b.dataUrl, brightness: b.brightness }));
    const legacy = await classifyBackgrounds(deduped.map(b => ({
        filename: b.filename, size: 0, dataUrl: b.dataUrl, brightness: b.brightness
    })));
    legacy.sequence = sequence;
    return legacy;
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
        const brightness = img.brightness != null ? img.brightness : await getImageBrightness(img.dataUrl);
        analyzed.push({ ...img, brightness });
    }

    // Sort by brightness
    analyzed.sort((a, b) => a.brightness - b.brightness);

    // Deduplicate by filename
    const unique = [];
    const seenFn = new Set();
    for (const a of analyzed) {
        if (!seenFn.has(a.filename)) {
            unique.push(a);
            seenFn.add(a.filename);
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
    // Strategy: a brand logo appears on MANY slides (or in the slide master).
    // A customer-logo / one-off graphic only appears on 1-2 slides. So we
    // count how often each image is referenced and only consider images that
    // appear on >=3 slides OR are in a slide master. This prevents picking
    // up customer logos like "Lemonade" from a template's customer section.

    const refCount = {};
    const addRefs = async (relsPath, weight) => {
        const f = zip.file(relsPath);
        if (!f) return;
        const xml = await f.async('string');
        for (const m of xml.matchAll(/Target="\.\.\/media\/([^"]+)"/g)) {
            refCount[m[1]] = (refCount[m[1]] || 0) + weight;
        }
    };

    // Slides
    for (let i = 1; i <= 100; i++) {
        await addRefs(`ppt/slides/_rels/slide${i}.xml.rels`, 1);
    }
    // Slide masters get a strong boost (template-level branding)
    for (let m = 1; m <= 5; m++) {
        await addRefs(`ppt/slideMasters/_rels/slideMaster${m}.xml.rels`, 5);
    }

    // Filter to images that look like logos AND appear in multiple places
    let bestLogo = null;
    let bestScore = -1;

    for (const [filename, count] of Object.entries(refCount)) {
        // Must appear on multiple slides or be in a master
        if (count < 3) continue;

        const mediaFile = zip.file('ppt/media/' + filename);
        if (!mediaFile) continue;
        const data = await mediaFile.async('arraybuffer');
        // Logos are small-to-medium sized
        if (data.byteLength > 200000 || data.byteLength < 800) continue;

        const dataUrl = arrayBufferToDataURL(data, getMime(filename));
        const dims = await getImageDimensions(dataUrl);
        if (!dims) continue;

        // Skip background-sized images
        if (dims.width > 1800 || dims.height > 800) continue;
        if (dims.width < 80 || dims.height < 20) continue;

        const aspect = dims.width / dims.height;
        if (aspect < 1.2 || aspect > 15) continue;

        const analysis = await analyzeLogoPixels(dataUrl);
        if (!analysis || analysis.visibleRatio < 0.03) continue;

        let score = count * 2;  // recurrence is the strongest signal
        if (aspect > 2 && aspect < 8) score += 3;
        if (analysis.visibleRatio > 0.05 && analysis.visibleRatio < 0.6) score += 2;

        if (score > bestScore) {
            bestScore = score;
            bestLogo = { dataUrl, isLight: analysis.isLight };
        }
    }

    if (bestLogo) {
        console.log(`Picked logo (score ${bestScore})`);
    } else {
        console.log('No reliable logo found in template — skipping logo overlay');
    }
    return bestLogo;
}

function analyzeLogoPixels(dataUrl) {
    return new Promise((resolve) => {
        const img = new Image();
        img.onload = () => {
            try {
                const w = Math.min(img.width, 200);
                const h = Math.round(img.height * (w / img.width));
                const canvas = document.createElement('canvas');
                canvas.width = w;
                canvas.height = h;
                const ctx = canvas.getContext('2d');
                ctx.drawImage(img, 0, 0, w, h);
                const data = ctx.getImageData(0, 0, w, h).data;
                let visible = 0;
                let lightSum = 0;
                let total = 0;
                for (let i = 0; i < data.length; i += 4) {
                    total++;
                    const a = data[i + 3];
                    if (a < 30) continue;
                    const r = data[i], g = data[i + 1], b = data[i + 2];
                    // Skip near-white pixels (background)
                    if (r > 240 && g > 240 && b > 240) continue;
                    visible++;
                    lightSum += (r + g + b) / 3;
                }
                const visibleRatio = visible / total;
                const avgLight = visible > 0 ? lightSum / visible : 0;
                resolve({ visibleRatio, isLight: avgLight > 180 });
            } catch (e) {
                resolve(null);
            }
        };
        img.onerror = () => resolve(null);
        img.src = dataUrl;
    });
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
        // Show logo on dark backdrop if it's a light/white logo, so it's visible in preview
        const bgStyle = brand.logoIsLight ? 'background:#0F161A;' : 'background:#FFFFFF;';
        logoBox.innerHTML = `<img src="${brand.logo}" alt="Logo" style="${bgStyle}padding:4px;">`;
        logoBox.dataset.logo = brand.logo;
        logoBox.dataset.logoLight = brand.logoIsLight ? '1' : '0';
    } else {
        logoBox.innerHTML = '<span class="no-logo-text">No logo found</span>';
        logoBox.dataset.logo = '';
        logoBox.dataset.logoLight = '0';
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

    // Store backgrounds in a JS variable (sequence may include many large data URLs)
    window.__previewBackgrounds = brand.backgrounds;
    brandPreview.classList.remove('hidden');
}

function getBrandFromPreview() {
    const name = brandNameInput.value.trim();
    if (!name) { alert('Please enter a brand name.'); return null; }
    const logoBox = $('#extracted-logo-box');
    const backgrounds = window.__previewBackgrounds || { title: null, dark: null, light: null, sequence: [] };

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
        logoIsLight: logoBox.dataset.logoLight === '1',
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
    const clearBtn = `<button class="clear-all-brands-btn" type="button">Clear all saved brands</button>`;
    savedBrandsList.innerHTML = clearBtn + brands.map((b, i) => `
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
    const clearAllBtn = savedBrandsList.querySelector('.clear-all-brands-btn');
    if (clearAllBtn) {
        clearAllBtn.addEventListener('click', () => {
            if (!confirm('Delete ALL saved brands? This cannot be undone.')) return;
            brands.length = 0;
            persistBrands();
            activeBrand = null;
            activeBrandBadge.textContent = 'None';
            window.__previewBackgrounds = null;
            brandPreview.classList.add('hidden');
            renderBrands();
        });
    }
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

    // Normalize line endings and strip leading/trailing blank lines
    const normalized = text.replace(/\r\n/g, '\n').trim();

    // STEP 1: Find slide boundaries.
    // A new slide starts when we see ANY of:
    //   - A "---" separator line
    //   - A line starting with "# " or "## "
    //   - A line matching "Slide N" / "Slide N —" / "Slide N:" / "Slide N." / "Slide N -"
    //   - A numbered line "1." / "1)" at start (only if it looks like a slide title, not a list item)
    const lines = normalized.split('\n');
    const slideStartIdx = [];
    const isSlideHeader = (line) => {
        const l = line.trim();
        if (!l) return false;
        if (l === '---' || l === '***') return 'sep';
        if (/^#{1,2}\s+/.test(l)) return 'hash';
        if (/^slide\s*\d+\b/i.test(l)) return 'slidekw';
        return false;
    };

    for (let i = 0; i < lines.length; i++) {
        if (isSlideHeader(lines[i])) slideStartIdx.push(i);
    }

    // If we found NO headers, fall back to splitting by blank lines
    let sections = [];
    if (slideStartIdx.length === 0) {
        sections = normalized.split(/\n\s*\n/).map(s => s.trim()).filter(Boolean);
    } else {
        // If first non-blank line isn't a header, treat it as an implicit first section
        const firstHeaderLine = slideStartIdx[0];
        if (firstHeaderLine > 0) {
            const pre = lines.slice(0, firstHeaderLine).join('\n').trim();
            if (pre) sections.push(pre);
        }
        for (let i = 0; i < slideStartIdx.length; i++) {
            const start = slideStartIdx[i];
            const end = i + 1 < slideStartIdx.length ? slideStartIdx[i + 1] : lines.length;
            const chunk = lines.slice(start, end).join('\n').trim();
            // Skip pure separator lines
            if (chunk === '---' || chunk === '***') continue;
            if (chunk) sections.push(chunk);
        }
    }

    // STEP 2: Convert each section into a slide
    const slides = [];
    for (let idx = 0; idx < sections.length; idx++) {
        const section = sections[idx];
        const sLines = section.split('\n').map(l => l).filter((l, i, arr) => !(i === 0 && !l.trim()));
        if (sLines.length === 0) continue;

        let firstLine = sLines[0].trim();
        let isTitleSlide = false;
        let heading = '';

        // Strip leading separator if present
        if (firstLine === '---' || firstLine === '***') {
            sLines.shift();
            firstLine = (sLines[0] || '').trim();
        }

        // Detect heading style
        if (/^#\s+/.test(firstLine) && !/^##\s+/.test(firstLine)) {
            heading = firstLine.replace(/^#\s+/, '').trim();
            isTitleSlide = true;
            sLines.shift();
        } else if (/^##\s+/.test(firstLine)) {
            heading = firstLine.replace(/^##\s+/, '').trim();
            sLines.shift();
        } else if (/^slide\s*\d+\b/i.test(firstLine)) {
            // "Slide 2 — What is Innovation?" → heading = "What is Innovation?"
            heading = firstLine.replace(/^slide\s*\d+\s*[—\-:.)]*\s*/i, '').trim();
            // First slide using this pattern is treated as title
            if (idx === 0) isTitleSlide = true;
            sLines.shift();
        } else {
            // No header marker — use first line as heading
            heading = firstLine;
            if (idx === 0) isTitleSlide = true;
            sLines.shift();
        }

        const bodyLines = sLines.map(l => l.trim()).filter(Boolean);
        const bodyText = bodyLines.join('\n');

        // Title / thank-you slide
        if (isTitleSlide) {
            slides.push({
                type: isThankYou(heading) ? 'thank-you' : 'title',
                bgType: 'title',
                data: { title: heading, subtitle: bodyText }
            });
            continue;
        }

        // Two columns
        if (/LEFT:/i.test(bodyText) && /RIGHT:/i.test(bodyText)) {
            const leftMatch = bodyText.match(/LEFT:\s*([\s\S]*?)(?=RIGHT:)/i);
            const rightMatch = bodyText.match(/RIGHT:\s*([\s\S]*)/i);
            slides.push({
                type: 'two-column', bgType: 'auto',
                data: {
                    heading,
                    left: leftMatch ? leftMatch[1].trim() : '',
                    right: rightMatch ? rightMatch[1].trim() : '',
                }
            });
            continue;
        }

        // Bullets — if MOST body lines start with - or * or •
        const bulletLines = bodyLines.filter(l => /^[-*•]\s+/.test(l));
        if (bulletLines.length >= 2 && bulletLines.length >= bodyLines.length * 0.6) {
            const bullets = bodyLines
                .map(l => l.replace(/^[-*•]\s+/, ''))
                .filter(Boolean)
                .join('\n');
            slides.push({
                type: 'bullets', bgType: 'auto',
                data: { heading, bullets }
            });
            continue;
        }

        // Regular content
        slides.push({
            type: 'content', bgType: 'auto',
            data: { heading, body: bodyText }
        });
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
    const sequence = (b.backgrounds && b.backgrounds.sequence) || [];

    let introHtml = '';
    let contentSlidesForPreview = parsedSlides;

    if (sequence.length > 0 && parsedSlides.length > 0) {
        // Cover
        const coverBg = sequence[0];
        const coverIsDark = coverBg.brightness < 140;
        const coverColor = coverIsDark ? b.textLight : b.textDark;
        introHtml += `
        <div class="preview-slide" style="background-image:url('${coverBg.dataUrl}');background-size:cover;background-position:center;">
            <span class="preview-slide-number">Cover</span>
            <div style="flex:1;display:flex;align-items:center;justify-content:center;">
                <div style="font-size:54px;font-weight:800;letter-spacing:2px;color:${coverColor};font-family:'${b.headingFont}',sans-serif;text-align:center;">${escapeHTML((b.name || 'Presentation').toUpperCase())}</div>
            </div>
        </div>`;

        // Title
        const titleBg = sequence[1] || sequence[0];
        const titleIsDark = titleBg.brightness < 140;
        const titleHeadingColor = titleIsDark ? b.textLight : b.primaryColor;
        const titleSubColor = titleIsDark ? b.textLight : b.textDark;
        const first = parsedSlides[0];
        const headlineText = first.data.title || first.data.heading || b.name || 'Presentation';
        const subtitleText = first.data.subtitle || first.data.body || '';
        introHtml += `
        <div class="preview-slide" style="background-image:url('${titleBg.dataUrl}');background-size:cover;background-position:center;">
            <span class="preview-slide-number">Title</span>
            <div style="flex:1;display:flex;flex-direction:column;justify-content:center;padding:0 20px;">
                <div style="font-size:32px;font-weight:800;color:${titleHeadingColor};font-family:'${b.headingFont}',sans-serif;line-height:1.15;">${escapeHTML(headlineText)}</div>
                ${subtitleText ? `<div style="font-size:14px;color:${titleSubColor};margin-top:14px;opacity:0.9;font-family:'${b.bodyFont}',sans-serif;">${escapeHTML(subtitleText)}</div>` : ''}
            </div>
        </div>`;

        // Agenda
        const agendaItems = parsedSlides.slice(1)
            .map(s => s.data.heading || s.data.title || '')
            .filter(Boolean);
        if (agendaItems.length > 0) {
            const agendaBg = sequence[2] || sequence[1] || sequence[0];
            const agendaIsDark = agendaBg.brightness < 140;
            const agendaHeading = agendaIsDark ? b.textLight : b.primaryColor;
            const agendaText = agendaIsDark ? b.textLight : b.textDark;
            const agendaList = agendaItems.map((t, i) =>
                `<li style="margin-bottom:6px;"><span style="opacity:0.6;font-weight:600;">${String(i + 1).padStart(2, '0')}.</span> ${escapeHTML(t)}</li>`
            ).join('');
            introHtml += `
            <div class="preview-slide" style="background-image:url('${agendaBg.dataUrl}');background-size:cover;background-position:center;">
                <span class="preview-slide-number">Agenda</span>
                <div style="font-size:22px;font-weight:700;color:${agendaHeading};font-family:'${b.headingFont}',sans-serif;margin-bottom:14px;">Today's Agenda</div>
                <ul style="list-style:none;padding:0;margin:0;font-size:14px;color:${agendaText};font-family:'${b.bodyFont}',sans-serif;">${agendaList}</ul>
            </div>`;
        }

        // First user slide is now used as the title slide, so skip it in content
        contentSlidesForPreview = parsedSlides.slice(1);
    }

    previewContainer.innerHTML = introHtml + contentSlidesForPreview.map((s, i) => {
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

    // Build the background sequence: cycle through ALL the template's slide
    // backgrounds in order, so the generated deck mirrors the template's flow.
    const sequence = (b.backgrounds && b.backgrounds.sequence) || [];
    const pickBg = (slideIndex, slideType) => {
        if (sequence.length === 0) {
            // Fall back to the legacy 3-bucket system
            const key = slideType === 'title' || slideType === 'thank-you' ? 'title' : 'auto';
            const fallback = b.backgrounds[key === 'auto' ? 'light' : key] || b.backgrounds.dark || b.backgrounds.title;
            return fallback ? { dataUrl: fallback, brightness: 128 } : null;
        }
        // Title slide always uses sequence[0] (first slide of template)
        if ((slideType === 'title' || slideType === 'thank-you') && sequence.length > 0) {
            return sequence[0];
        }
        // Content slides cycle through sequence[1..] in order
        if (sequence.length === 1) return sequence[0];
        const contentSeq = sequence.slice(1);
        return contentSeq[(slideIndex - 1) % contentSeq.length];
    };

    // Helper to set a slide background and return its brightness/colors
    const applyBg = (slide, bgEntry) => {
        if (!bgEntry) {
            slide.background = { color: '0F161A' };
            return { isDark: true, textColor: hex(b.textLight), headingColor: hex(b.textLight) };
        }
        const bgData = bgEntry.dataUrl.startsWith('data:') ? bgEntry.dataUrl.substring(5) : bgEntry.dataUrl;
        slide.background = { data: bgData };
        const isDark = bgEntry.brightness < 140;
        return {
            isDark,
            textColor: isDark ? hex(b.textLight) : hex(b.textDark),
            headingColor: isDark ? hex(b.textLight) : hex(b.primaryColor),
        };
    };

    // === AUTO-PREPENDED INTRO SLIDES ===
    // 1. Cover (brand name on template's first background)
    // 2. Title (user's first slide as big headline + tagline)
    // 3. Agenda (auto-built from headings of remaining content slides)
    if (sequence.length > 0) {
        // -- Slide 1: COVER --
        const cover = pptx.addSlide();
        const coverColors = applyBg(cover, sequence[0]);
        cover.addText((b.name || 'Presentation').toUpperCase(), {
            x: 0.5, y: 2.6, w: 12.3, h: 2.2,
            fontSize: 96, fontFace: b.headingFont,
            color: coverColors.textColor, bold: true,
            align: 'center', valign: 'middle',
        });

        // -- Slide 2: TITLE -- (uses user's first slide as headline + subtitle)
        const titleSlide = pptx.addSlide();
        const titleBg = sequence[1] || sequence[0];
        const titleColors = applyBg(titleSlide, titleBg);
        const firstUserSlide = parsedSlides[0];
        const headlineText = firstUserSlide ? (firstUserSlide.data.title || firstUserSlide.data.heading || b.name) : (b.name || 'Presentation');
        const subtitleText = firstUserSlide ? (firstUserSlide.data.subtitle || firstUserSlide.data.body || '') : '';
        titleSlide.addText(headlineText, {
            x: 0.6, y: 1.6, w: 12.1, h: 3.2,
            fontSize: 60, fontFace: b.headingFont,
            color: titleColors.headingColor, bold: true,
            align: 'left', valign: 'middle',
        });
        if (subtitleText) {
            titleSlide.addText(subtitleText, {
                x: 0.6, y: 5.0, w: 12.1, h: 1.6,
                fontSize: 18, fontFace: b.bodyFont,
                color: titleColors.textColor,
                align: 'left', valign: 'top',
            });
        }

        // -- Slide 3: AGENDA -- (auto-built from headings of all content slides after the first)
        const agendaItems = parsedSlides.slice(1)
            .map(s => s.data.heading || s.data.title || '')
            .filter(Boolean);
        if (agendaItems.length > 0) {
            const agendaSlide = pptx.addSlide();
            const agendaBg = sequence[2] || sequence[1] || sequence[0];
            const agendaColors = applyBg(agendaSlide, agendaBg);
            agendaSlide.addText("Today's Agenda", {
                x: 0.6, y: 0.6, w: 12.1, h: 1.0,
                fontSize: 36, fontFace: b.headingFont,
                color: agendaColors.headingColor, bold: true,
            });
            const numbered = agendaItems.map((text, i) => ({
                text: `${String(i + 1).padStart(2, '0')}.  ${text}`,
                options: { fontSize: 22, fontFace: b.bodyFont, color: agendaColors.textColor, breakLine: true },
            }));
            agendaSlide.addText(numbered, {
                x: 0.8, y: 1.8, w: 11.7, h: 5.0,
                valign: 'top',
            });
        }
    }

    // The first user slide has been used as the title slide above, so we
    // skip it in the main content loop. Everything else renders as normal.
    const contentSlides = sequence.length > 0 && parsedSlides.length > 0
        ? parsedSlides.slice(1)
        : parsedSlides;

    // Content slides cycle through sequence starting AFTER cover/title/agenda
    // (cover used sequence[0], title used sequence[1], agenda used sequence[2])
    const contentBgPool = sequence.length > 3 ? sequence.slice(3) : (sequence.length > 1 ? sequence.slice(1) : sequence);

    contentSlides.forEach((s, idx) => {
        const slide = pptx.addSlide();

        // Pick background by cycling through the content pool
        const bgEntry = contentBgPool.length > 0 ? contentBgPool[idx % contentBgPool.length] : null;
        const colors = applyBg(slide, bgEntry);
        const isDark = colors.isDark;
        const textColor = colors.textColor;
        const headingColor = colors.headingColor;

        // Logo: only render if it will be visible against the background
        // (white logo on dark bg = visible; white logo on light bg = invisible, skip)
        if (b.logo) {
            const logoVisible = b.logoIsLight ? isDark : !isDark;
            if (logoVisible) {
                slide.addImage({
                    data: b.logo,
                    x: 0.4, y: 6.8,
                    w: 1.2, h: 0.4,
                    sizing: { type: 'contain', w: 1.2, h: 0.4 },
                });
            }
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
