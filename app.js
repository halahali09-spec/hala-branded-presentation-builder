// ============================================================
// BRANDED PRESENTATION BUILDER — app.js
// Upload a branded .pptx → paste content → download branded presentation
// ============================================================

// ---- State ----
let brands = JSON.parse(localStorage.getItem('pptx-brands') || '[]');
let activeBrand = null;
let parsedSlides = [];

// ---- DOM refs ----
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

// Steps
const stepEls = $$('.step');
const panels = { 1: $('#step1'), 2: $('#step2'), 3: $('#step3') };

// Step 1
const dropZone = $('#drop-zone');
const pptxUpload = $('#pptx-upload');
const extractStatus = $('#extract-status');
const brandPreview = $('#brand-preview');
const brandNameInput = $('#brand-name');
const savedBrandsList = $('#saved-brands-list');
const saveBrandBtn = $('#save-brand-btn');
const useBrandBtn = $('#use-brand-btn');
const manualLogo = $('#manual-logo');

// Step 2
const contentInput = $('#content-input');
const parsedPreviewEl = $('#parsed-preview');
const generateBtn = $('#generate-btn');
const activeBrandBadge = $('#active-brand-badge');
const formatGuideToggle = $('#format-guide-toggle');
const formatGuideBody = $('#format-guide-body');

// Step 3
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
// STEP 1 — BRAND EXTRACTION FROM .PPTX
// ============================================================

// -- Drop zone events --
dropZone.addEventListener('click', () => pptxUpload.click());

dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
});

dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('drag-over');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file && file.name.endsWith('.pptx')) {
        handlePptxUpload(file);
    }
});

pptxUpload.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) handlePptxUpload(file);
});

// -- Manual logo upload --
manualLogo.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const dataUrl = await fileToDataURL(file);
    const logoBox = $('#extracted-logo-box');
    logoBox.innerHTML = `<img src="${dataUrl}" alt="Logo">`;
    logoBox.dataset.logo = dataUrl;
});

// -- Extract branding from uploaded .pptx --
async function handlePptxUpload(file) {
    extractStatus.classList.remove('hidden');
    brandPreview.classList.add('hidden');

    try {
        const brand = await extractBrandFromPptx(file);
        showBrandPreview(brand);
    } catch (err) {
        console.error('Extraction error:', err);
        alert('Could not read branding from this file. Please try a different .pptx file.');
    }

    extractStatus.classList.add('hidden');
}

async function extractBrandFromPptx(file) {
    const arrayBuffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);

    const brand = {
        name: file.name.replace('.pptx', '').replace(/[_-]/g, ' '),
        primaryColor: '#2563EB',
        secondaryColor: '#1E40AF',
        accentColor: '#F59E0B',
        textColor: '#1F2937',
        bgColor: '#FFFFFF',
        headingFont: 'Calibri Light',
        bodyFont: 'Calibri',
        logo: null,
    };

    // 1. Extract theme colors and fonts
    const themeFile = zip.file('ppt/theme/theme1.xml');
    if (themeFile) {
        const themeXml = await themeFile.async('string');
        const parser = new DOMParser();
        const doc = parser.parseFromString(themeXml, 'application/xml');

        // Extract colors
        const colors = extractThemeColors(doc, themeXml);
        if (colors.accent1) brand.primaryColor = colors.accent1;
        if (colors.accent2) brand.secondaryColor = colors.accent2;
        if (colors.accent3) brand.accentColor = colors.accent3;
        if (colors.dk1) brand.textColor = colors.dk1;
        if (colors.lt1) brand.bgColor = colors.lt1;

        // Extract fonts
        const fonts = extractThemeFonts(doc, themeXml);
        if (fonts.major) brand.headingFont = fonts.major;
        if (fonts.minor) brand.bodyFont = fonts.minor;
    }

    // 2. Extract logo from slide master
    const logo = await extractLogo(zip);
    if (logo) brand.logo = logo;

    return brand;
}

function extractThemeColors(doc, xmlString) {
    const colors = {};
    const colorNames = ['dk1', 'dk2', 'lt1', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];

    // Try namespace-aware approach first
    const NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    let clrScheme = doc.getElementsByTagNameNS(NS, 'clrScheme')[0];

    if (clrScheme) {
        for (const child of clrScheme.children) {
            const name = child.localName;
            if (colorNames.includes(name)) {
                const hex = getColorFromNode(child);
                if (hex) colors[name] = '#' + hex;
            }
        }
    }

    // Fallback: regex extraction from raw XML
    if (Object.keys(colors).length === 0) {
        for (const name of colorNames) {
            // Match <a:accent1><a:srgbClr val="RRGGBB"/></a:accent1>
            const srgbMatch = xmlString.match(new RegExp(`<a:${name}>\\s*<a:srgbClr val="([A-Fa-f0-9]{6})"`, 'i'));
            if (srgbMatch) {
                colors[name] = '#' + srgbMatch[1];
                continue;
            }
            // Match <a:dk1><a:sysClr ... lastClr="RRGGBB"/></a:dk1>
            const sysMatch = xmlString.match(new RegExp(`<a:${name}>\\s*<a:sysClr[^>]*lastClr="([A-Fa-f0-9]{6})"`, 'i'));
            if (sysMatch) {
                colors[name] = '#' + sysMatch[1];
            }
        }
    }

    return colors;
}

function getColorFromNode(parentNode) {
    for (const child of parentNode.children) {
        if (child.localName === 'srgbClr') {
            return child.getAttribute('val');
        }
        if (child.localName === 'sysClr') {
            return child.getAttribute('lastClr');
        }
    }
    return null;
}

function extractThemeFonts(doc, xmlString) {
    const fonts = {};
    const NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';

    // Try namespace-aware
    const majorFonts = doc.getElementsByTagNameNS(NS, 'majorFont');
    const minorFonts = doc.getElementsByTagNameNS(NS, 'minorFont');

    if (majorFonts.length > 0) {
        const latin = majorFonts[0].getElementsByTagNameNS(NS, 'latin')[0];
        if (latin) fonts.major = latin.getAttribute('typeface');
    }
    if (minorFonts.length > 0) {
        const latin = minorFonts[0].getElementsByTagNameNS(NS, 'latin')[0];
        if (latin) fonts.minor = latin.getAttribute('typeface');
    }

    // Fallback: regex
    if (!fonts.major) {
        const majorMatch = xmlString.match(/<a:majorFont>[\s\S]*?<a:latin typeface="([^"]+)"/i);
        if (majorMatch) fonts.major = majorMatch[1];
    }
    if (!fonts.minor) {
        const minorMatch = xmlString.match(/<a:minorFont>[\s\S]*?<a:latin typeface="([^"]+)"/i);
        if (minorMatch) fonts.minor = minorMatch[1];
    }

    return fonts;
}

async function extractLogo(zip) {
    // Strategy: find images referenced in the slide master
    const masterRelsFile = zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels');
    if (!masterRelsFile) return null;

    const relsXml = await masterRelsFile.async('string');
    // Find all image relationships
    const imageRefs = [];
    const regex = /Target="\.\.\/media\/([^"]+)"/g;
    let match;
    while ((match = regex.exec(relsXml)) !== null) {
        imageRefs.push(match[1]);
    }

    if (imageRefs.length === 0) {
        // Try slide layouts too
        return await extractLogoFromLayouts(zip);
    }

    // Pick the best logo candidate (smallest file = likely a logo, not a background)
    let bestLogo = null;
    let smallestSize = Infinity;

    for (const filename of imageRefs) {
        const mediaFile = zip.file('ppt/media/' + filename);
        if (!mediaFile) continue;

        const data = await mediaFile.async('arraybuffer');
        if (data.byteLength < smallestSize && data.byteLength > 100) {
            smallestSize = data.byteLength;
            const ext = filename.split('.').pop().toLowerCase();
            const mimeMap = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', svg: 'image/svg+xml' };
            const mime = mimeMap[ext] || 'image/png';
            if (mime.startsWith('image/')) {
                bestLogo = arrayBufferToDataURL(data, mime);
            }
        }
    }

    return bestLogo;
}

async function extractLogoFromLayouts(zip) {
    // Check first few slide layouts for images
    for (let i = 1; i <= 3; i++) {
        const relsFile = zip.file(`ppt/slideLayouts/_rels/slideLayout${i}.xml.rels`);
        if (!relsFile) continue;

        const relsXml = await relsFile.async('string');
        const match = relsXml.match(/Target="\.\.\/media\/([^"]+)"/);
        if (match) {
            const mediaFile = zip.file('ppt/media/' + match[1]);
            if (mediaFile) {
                const data = await mediaFile.async('arraybuffer');
                const ext = match[1].split('.').pop().toLowerCase();
                const mimeMap = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif' };
                const mime = mimeMap[ext] || 'image/png';
                return arrayBufferToDataURL(data, mime);
            }
        }
    }
    return null;
}

function arrayBufferToDataURL(buffer, mime) {
    const bytes = new Uint8Array(buffer);
    let binary = '';
    for (let i = 0; i < bytes.length; i++) {
        binary += String.fromCharCode(bytes[i]);
    }
    return 'data:' + mime + ';base64,' + btoa(binary);
}

// -- Show extracted brand in the UI --
function showBrandPreview(brand) {
    $('#ex-primary').value = brand.primaryColor;
    $('#ex-secondary').value = brand.secondaryColor;
    $('#ex-accent').value = brand.accentColor;
    $('#ex-text').value = brand.textColor;
    $('#ex-bg').value = brand.bgColor;
    $('#ex-heading-font').value = brand.headingFont;
    $('#ex-body-font').value = brand.bodyFont;
    brandNameInput.value = brand.name;

    const logoBox = $('#extracted-logo-box');
    if (brand.logo) {
        logoBox.innerHTML = `<img src="${brand.logo}" alt="Logo">`;
        logoBox.dataset.logo = brand.logo;
    } else {
        logoBox.innerHTML = '<span class="no-logo-text">No logo found</span>';
        logoBox.dataset.logo = '';
    }

    brandPreview.classList.remove('hidden');
}

function getBrandFromPreview() {
    const name = brandNameInput.value.trim();
    if (!name) { alert('Please enter a brand name.'); return null; }
    const logoBox = $('#extracted-logo-box');
    return {
        name,
        primaryColor: $('#ex-primary').value,
        secondaryColor: $('#ex-secondary').value,
        accentColor: $('#ex-accent').value,
        textColor: $('#ex-text').value,
        bgColor: $('#ex-bg').value,
        headingFont: $('#ex-heading-font').value || 'Calibri',
        bodyFont: $('#ex-body-font').value || 'Calibri',
        logo: logoBox.dataset.logo || null,
    };
}

// -- Save / Use brand --
saveBrandBtn.addEventListener('click', () => {
    const b = getBrandFromPreview();
    if (!b) return;
    const existingIdx = brands.findIndex(x => x.name.toLowerCase() === b.name.toLowerCase());
    if (existingIdx >= 0) {
        brands[existingIdx] = b;
    } else {
        brands.push(b);
    }
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

// -- Render saved brands --
function renderBrands() {
    if (brands.length === 0) {
        savedBrandsList.innerHTML = '<p class="no-brands">No saved brands yet. Upload a .pptx to get started!</p>';
        return;
    }
    savedBrandsList.innerHTML = brands.map((b, i) => `
        <div class="brand-card" data-index="${i}">
            <button class="brand-card-delete" data-index="${i}" title="Delete brand">&times;</button>
            <div class="brand-card-name">${escapeHTML(b.name)}</div>
            <div class="brand-card-colors">
                <div class="brand-card-swatch" style="background:${b.primaryColor}"></div>
                <div class="brand-card-swatch" style="background:${b.secondaryColor}"></div>
                <div class="brand-card-swatch" style="background:${b.accentColor}"></div>
            </div>
            ${b.logo ? `<img class="brand-card-logo" src="${b.logo}" alt="Logo">` : ''}
        </div>
    `).join('');

    // Click to load brand
    savedBrandsList.querySelectorAll('.brand-card').forEach(card => {
        card.addEventListener('click', (e) => {
            if (e.target.classList.contains('brand-card-delete')) return;
            const b = brands[parseInt(card.dataset.index)];
            showBrandPreview(b);
        });
    });

    // Delete
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

// Format guide toggle
formatGuideToggle.addEventListener('click', () => {
    formatGuideBody.classList.toggle('hidden');
    formatGuideToggle.classList.toggle('open');
});

// Content input — parse as user types
contentInput.addEventListener('input', () => {
    parsedSlides = parseContent(contentInput.value);
    renderParsedPreview();
    generateBtn.disabled = parsedSlides.length === 0;
});

function parseContent(text) {
    if (!text.trim()) return [];

    // Split by --- separator
    const sections = text.split(/\n---\n|\n---$|^---\n/);
    const slides = [];

    for (const section of sections) {
        const trimmed = section.trim();
        if (!trimmed) continue;

        const lines = trimmed.split('\n');
        const firstLine = lines[0].trim();

        // Title slide: starts with single #
        if (firstLine.startsWith('# ') && !firstLine.startsWith('## ')) {
            const title = firstLine.replace(/^# /, '');
            const subtitle = lines.slice(1).map(l => l.trim()).filter(l => l).join('\n');
            slides.push({
                type: isThankYou(title) ? 'thank-you' : 'title',
                data: { title, subtitle }
            });
            continue;
        }

        // Content/bullet slides: starts with ##
        if (firstLine.startsWith('## ')) {
            const heading = firstLine.replace(/^## /, '');
            const bodyLines = lines.slice(1).map(l => l.trim()).filter(l => l);

            // Check for two-column format
            const bodyText = bodyLines.join('\n');
            if (bodyText.includes('LEFT:') && bodyText.includes('RIGHT:')) {
                const leftMatch = bodyText.match(/LEFT:\s*([\s\S]*?)(?=RIGHT:)/i);
                const rightMatch = bodyText.match(/RIGHT:\s*([\s\S]*)/i);
                slides.push({
                    type: 'two-column',
                    data: {
                        heading,
                        left: leftMatch ? leftMatch[1].trim() : '',
                        right: rightMatch ? rightMatch[1].trim() : '',
                    }
                });
                continue;
            }

            // Check if all body lines are bullets
            const allBullets = bodyLines.length > 0 && bodyLines.every(l => l.startsWith('- ') || l.startsWith('* '));
            if (allBullets) {
                slides.push({
                    type: 'bullets',
                    data: {
                        heading,
                        bullets: bodyLines.map(l => l.replace(/^[-*] /, '')).join('\n'),
                    }
                });
            } else {
                slides.push({
                    type: 'content',
                    data: {
                        heading,
                        body: bodyLines.join('\n'),
                    }
                });
            }
            continue;
        }

        // Plain text without heading — treat as content slide
        slides.push({
            type: 'content',
            data: {
                heading: '',
                body: trimmed,
            }
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

    const typeLabels = {
        'title': 'Title Slide',
        'thank-you': 'Thank You',
        'content': 'Content',
        'bullets': 'Bullets',
        'two-column': 'Two Columns',
    };

    parsedPreviewEl.innerHTML = `<p style="font-size:13px;font-weight:600;color:#475569;margin-bottom:8px;">${parsedSlides.length} slide${parsedSlides.length > 1 ? 's' : ''} detected:</p>` +
        parsedSlides.map((s, i) => {
            const title = s.data.title || s.data.heading || 'Untitled';
            return `
            <div class="parsed-slide-card">
                <div class="parsed-slide-num">${i + 1}</div>
                <div>
                    <div class="parsed-slide-type">${typeLabels[s.type] || s.type}</div>
                    <div class="parsed-slide-title">${escapeHTML(title)}</div>
                </div>
            </div>`;
        }).join('');
}

// Generate button
generateBtn.addEventListener('click', () => {
    parsedSlides = parseContent(contentInput.value);
    if (parsedSlides.length > 0) {
        goToStep(3);
    }
});

// ============================================================
// STEP 3 — PREVIEW & DOWNLOAD
// ============================================================

function renderPreview() {
    const b = activeBrand;
    previewContainer.innerHTML = parsedSlides.map((s, i) => {
        let inner = '';

        if (s.type === 'title' || s.type === 'thank-you') {
            inner = `
                <div style="flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;">
                    <div style="font-size:28px;font-weight:700;color:${b.primaryColor};font-family:'${b.headingFont}',sans-serif;">${escapeHTML(s.data.title || '')}</div>
                    <div style="font-size:16px;color:${b.textColor};margin-top:12px;font-family:'${b.bodyFont}',sans-serif;">${escapeHTML(s.data.subtitle || '')}</div>
                </div>`;
        } else if (s.type === 'content') {
            inner = `
                <div style="font-size:22px;font-weight:700;color:${b.primaryColor};font-family:'${b.headingFont}',sans-serif;margin-bottom:16px;">${escapeHTML(s.data.heading || '')}</div>
                <div style="font-size:14px;color:${b.textColor};font-family:'${b.bodyFont}',sans-serif;white-space:pre-wrap;flex:1;">${escapeHTML(s.data.body || '')}</div>`;
        } else if (s.type === 'bullets') {
            const bullets = (s.data.bullets || '').split('\n').filter(x => x.trim()).map(x => `<li>${escapeHTML(x)}</li>`).join('');
            inner = `
                <div style="font-size:22px;font-weight:700;color:${b.primaryColor};font-family:'${b.headingFont}',sans-serif;margin-bottom:16px;">${escapeHTML(s.data.heading || '')}</div>
                <ul style="font-size:14px;color:${b.textColor};font-family:'${b.bodyFont}',sans-serif;padding-left:24px;flex:1;">${bullets}</ul>`;
        } else if (s.type === 'two-column') {
            inner = `
                <div style="font-size:22px;font-weight:700;color:${b.primaryColor};font-family:'${b.headingFont}',sans-serif;margin-bottom:16px;">${escapeHTML(s.data.heading || '')}</div>
                <div style="display:flex;gap:24px;flex:1;">
                    <div style="flex:1;font-size:13px;color:${b.textColor};font-family:'${b.bodyFont}',sans-serif;white-space:pre-wrap;">${escapeHTML(s.data.left || '')}</div>
                    <div style="flex:1;font-size:13px;color:${b.textColor};font-family:'${b.bodyFont}',sans-serif;white-space:pre-wrap;">${escapeHTML(s.data.right || '')}</div>
                </div>`;
        }

        const logoTag = b.logo ? `<img class="preview-slide-logo" src="${b.logo}">` : '';

        return `
        <div class="preview-slide" style="background:${b.bgColor};font-family:'${b.bodyFont}',sans-serif;border-top:6px solid ${b.primaryColor};">
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

    const hex = (c) => c.replace('#', '');

    parsedSlides.forEach((s) => {
        const slide = pptx.addSlide();
        slide.background = { color: hex(b.bgColor) };

        // Top accent bar
        slide.addShape(pptx.ShapeType.rect, {
            x: 0, y: 0, w: '100%', h: 0.12,
            fill: { color: hex(b.primaryColor) },
        });

        // Logo on every slide
        if (b.logo) {
            slide.addImage({
                data: b.logo,
                x: 11.2, y: 6.6,
                w: 1.5, h: 0.6,
                sizing: { type: 'contain', w: 1.5, h: 0.6 },
            });
        }

        if (s.type === 'title' || s.type === 'thank-you') {
            slide.addText(s.data.title || '', {
                x: 0.8, y: 2.0, w: 11.5, h: 1.5,
                fontSize: 36, fontFace: b.headingFont,
                color: hex(b.primaryColor),
                bold: true, align: 'center', valign: 'middle',
            });
            slide.addText(s.data.subtitle || '', {
                x: 0.8, y: 3.5, w: 11.5, h: 1.0,
                fontSize: 20, fontFace: b.bodyFont,
                color: hex(b.textColor),
                align: 'center', valign: 'top',
            });
        } else if (s.type === 'content') {
            slide.addText(s.data.heading || '', {
                x: 0.8, y: 0.4, w: 11.5, h: 0.8,
                fontSize: 28, fontFace: b.headingFont,
                color: hex(b.primaryColor),
                bold: true,
            });
            slide.addText(s.data.body || '', {
                x: 0.8, y: 1.4, w: 11.5, h: 5.2,
                fontSize: 16, fontFace: b.bodyFont,
                color: hex(b.textColor),
                valign: 'top',
            });
        } else if (s.type === 'bullets') {
            slide.addText(s.data.heading || '', {
                x: 0.8, y: 0.4, w: 11.5, h: 0.8,
                fontSize: 28, fontFace: b.headingFont,
                color: hex(b.primaryColor),
                bold: true,
            });
            const bulletItems = (s.data.bullets || '').split('\n').filter(x => x.trim()).map(text => ({
                text,
                options: { bullet: { type: 'bullet' }, fontSize: 16, fontFace: b.bodyFont, color: hex(b.textColor) },
            }));
            if (bulletItems.length > 0) {
                slide.addText(bulletItems, {
                    x: 0.8, y: 1.4, w: 11.5, h: 5.2,
                    valign: 'top',
                });
            }
        } else if (s.type === 'two-column') {
            slide.addText(s.data.heading || '', {
                x: 0.8, y: 0.4, w: 11.5, h: 0.8,
                fontSize: 28, fontFace: b.headingFont,
                color: hex(b.primaryColor),
                bold: true,
            });
            slide.addText(s.data.left || '', {
                x: 0.8, y: 1.4, w: 5.4, h: 5.2,
                fontSize: 15, fontFace: b.bodyFont,
                color: hex(b.textColor),
                valign: 'top',
            });
            slide.addShape(pptx.ShapeType.line, {
                x: 6.5, y: 1.4, w: 0, h: 5.0,
                line: { color: hex(b.accentColor), width: 1 },
            });
            slide.addText(s.data.right || '', {
                x: 6.9, y: 1.4, w: 5.4, h: 5.2,
                fontSize: 15, fontFace: b.bodyFont,
                color: hex(b.textColor),
                valign: 'top',
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
