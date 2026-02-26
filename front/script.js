class DocumentTypography {
    constructor() {
        this.docId = null;
        this.content = [];
        this.styles = [];
        this.history = [];
        this.redoStack = [];
        this.currentTool = null;
        this.savedSelection = null;
        this.pendingIconData = null;
        this.textColor = '#E53935';
        this.bgColor = '#FFEB3B';
        this.borderColor = '#000000';
        this.fontSize = '16px';
        this.letterSpacing = '4px';
        this.apiBase = '/api';
        this.COLOR_HISTORY_MAX = 6;
        this.colorHistoryKeys = { text: 'colorHistory_text', bg: 'colorHistory_bg', border: 'colorHistory_border' };
        this.colorHistory = { text: this.loadColorHistory('text'), bg: this.loadColorHistory('bg'), border: this.loadColorHistory('border') };
        this.initElements();
        this.initEventListeners();
        this.renderColorSwatches('text');
        this.renderColorSwatches('bg');
        this.renderColorSwatches('border');
    }

    initElements() {
        'fileInput uploadBtn uploadBtnAlt documentViewport documentContainer documentContent emptyState fileInfo selectionHint pickerText pickerBg pickerBorder highlightIcon textcolorIcon fontSizeInput fontSizeMinus fontSizePlus letterSpacingInput letterSpacingMinus letterSpacingPlus letterSpacingBtn letterSpacingPopover undoBtn redoBtn clearBtn saveBtn iconUploadBtn iconModal iconModalClose iconDescription iconModalCancel iconModalSubmit iconModalSubmitText iconModalSpinner stylesList styleCount toastContainer fontFamilyBtn fontFamilyPopover scriptSizeBtn scriptSizePopover swatchesText swatchesBg swatchesBorder previewText previewBg previewBorder floatingToolbar ftFontFamilyPopover ftScriptSizePopover ftFontSizeInput ftFontSizeMinus ftFontSizePlus'.split(' ').forEach(id => this[id] = document.getElementById(id));
        this.toolButtons = document.querySelectorAll('.tool-btn');
        this.fontOptions = document.querySelectorAll('.font-option');
        this.scriptOptions = document.querySelectorAll('.script-option');
        this.colorPopovers = { text: document.getElementById('colorPopoverText'), bg: document.getElementById('colorPopoverBg'), border: document.getElementById('colorPopoverBorder') };
        this.colorTriggers = document.querySelectorAll('.color-trigger');
    }

    initEventListeners() {
        this.uploadBtn.addEventListener('click', () => this.fileInput.click());
        this.uploadBtnAlt.addEventListener('click', () => this.fileInput.click());
        this.fileInput.addEventListener('change', e => this.handleFileUpload(e));
        this.documentViewport.addEventListener('dragover', e => { e.preventDefault(); this.documentViewport.classList.add('drag-over'); });
        this.documentViewport.addEventListener('dragleave', e => { e.preventDefault(); this.documentViewport.classList.remove('drag-over'); });
        this.documentViewport.addEventListener('drop', e => this.handleDrop(e));
        this.toolButtons.forEach(btn => { if (btn.dataset.tool) btn.addEventListener('click', () => this.selectTool(btn.dataset.tool)); });
        this.pickerText.addEventListener('input', e => { this.textColor = e.target.value; this.textcolorIcon.style.color = e.target.value; this.previewText.style.backgroundColor = e.target.value; });
        this.pickerBg.addEventListener('input', e => { this.bgColor = e.target.value; this.highlightIcon.style.background = e.target.value; this.previewBg.style.backgroundColor = e.target.value; });
        this.pickerBorder.addEventListener('input', e => { this.borderColor = e.target.value; this.previewBorder.style.backgroundColor = e.target.value; });
        this.colorTriggers.forEach(trigger => trigger.addEventListener('click', e => { e.stopPropagation(); this.toggleColorPopover(trigger.dataset.picker); }));
        document.addEventListener('click', e => { if (!e.target.closest('.color-picker-item')) this.closeAllColorPopovers(); });
        this._setupStepper(this.fontSizeMinus, this.fontSizePlus, this.fontSizeInput, 1, 200, val => { this.fontSize = val + 'px'; if (this.savedSelection) this.applyToolToSelection('fontsize', true); });
        this.fontSizeInput.addEventListener('change', e => {
            const val = e.target.value;
            if (val && parseInt(val) > 0) {
                this.fontSize = val + 'px';
                if (this.savedSelection) this.applyToolToSelection('fontsize', true);
            }
        });
        this._setupStepper(this.letterSpacingMinus, this.letterSpacingPlus, this.letterSpacingInput, 0, 100, val => { this.letterSpacing = val + 'px'; if (this.savedSelection) this.applyToolToSelection('letterspacing', true); });
        this.letterSpacingInput.addEventListener('change', e => {
            const val = e.target.value;
            if (val && parseInt(val) >= 0) {
                this.letterSpacing = val + 'px';
                if (this.savedSelection) this.applyToolToSelection('letterspacing', true);
            }
        });
        this.letterSpacingBtn.addEventListener('click', () => this.toggleLetterSpacingPopover());
        this.undoBtn.addEventListener('click', () => this.undo());
        this.redoBtn.addEventListener('click', () => this.redo());
        this.clearBtn.addEventListener('click', () => this.clearAllStyles());
        this.saveBtn.addEventListener('click', () => this.saveStyles());
        this.iconUploadBtn.addEventListener('click', () => {
            if (!this.docId) { this.showToast('Load a document first', 'error'); return; }
            this.openIconModal();
        });
        this.documentContent.addEventListener('click', e => this.handleIconPlacement(e));
        this.iconModalClose.addEventListener('click', () => this.closeIconModal());
        this.iconModalCancel.addEventListener('click', () => this.closeIconModal());
        this.iconModalSubmit.addEventListener('click', () => this.generateIcon());
        this.iconDescription.addEventListener('keydown', e => { if (e.key === 'Enter') this.generateIcon(); });
        this.iconModal.addEventListener('click', e => { if (e.target === this.iconModal) this.closeIconModal(); });
        this.fontFamilyBtn.addEventListener('click', () => this.toggleFontFamilyPopover());
        this.fontOptions.forEach(opt => opt.addEventListener('click', () => this.selectFontFamily(opt.dataset.font)));
        this.scriptSizeBtn.addEventListener('click', () => this.toggleScriptSizePopover());
        this.scriptOptions.forEach(opt => opt.addEventListener('click', () => this.selectScriptSize(opt.dataset.script)));
        document.addEventListener('mouseup', e => this.handleTextSelection(e));
        document.addEventListener('keydown', e => this.handleKeyboard(e));

        this.floatingToolbar.querySelectorAll('.ft-btn').forEach(btn => {
            btn.addEventListener('mousedown', e => e.preventDefault());
            btn.addEventListener('click', e => {
                e.preventDefault();
                e.stopPropagation();
                const tool = btn.dataset.ftTool;
                if (tool === 'fontfamily') { this.toggleFtPopover('ftFontFamilyPopover'); return; }
                if (tool === 'scriptsize') { this.toggleFtPopover('ftScriptSizePopover'); return; }
                if (tool === 'letterspacing') {
                    this.closeFtPopovers();
                    this.applyToolToSelection('letterspacing');
                    this.hideFloatingToolbar();
                    return;
                }
                this.closeFtPopovers();
                this.applyToolToSelection(tool);
                this.hideFloatingToolbar();
            });
        });
        this.ftFontFamilyPopover.querySelectorAll('[data-ft-font]').forEach(opt => {
            opt.addEventListener('mousedown', e => e.preventDefault());
            opt.addEventListener('click', e => {
                e.stopPropagation();
                this.applyToolToSelection(opt.dataset.ftFont);
                this.closeFtPopovers();
                this.hideFloatingToolbar();
            });
        });
        this.ftScriptSizePopover.querySelectorAll('[data-ft-script]').forEach(opt => {
            opt.addEventListener('mousedown', e => e.preventDefault());
            opt.addEventListener('click', e => {
                e.stopPropagation();
                this.applyToolToSelection(opt.dataset.ftScript);
                this.closeFtPopovers();
                this.hideFloatingToolbar();
            });
        });
        this.floatingToolbar.addEventListener('mousedown', e => e.preventDefault());

        this._setupStepper(this.ftFontSizeMinus, this.ftFontSizePlus, this.ftFontSizeInput, 1, 200, val => {
            this.fontSize = val + 'px';
            this.fontSizeInput.value = val;
            if (this.savedSelection) this.applyToolToSelection('fontsize', true);
        });
        this.ftFontSizeInput.addEventListener('change', e => {
            const val = e.target.value;
            if (val && parseInt(val) > 0) {
                this.fontSize = val + 'px';
                this.fontSizeInput.value = val;
                if (this.savedSelection) this.applyToolToSelection('fontsize', true);
            }
        });
    }

    async handleFileUpload(e) {
        const file = e.target.files[0];
        if (!file) return;
        const validTypes = ['application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'application/msword'];
        if (!validTypes.includes(file.type) && !file.name.endsWith('.docx')) {
            this.showToast('Please select a Word document (.docx)', 'error');
            return;
        }
        await this.uploadDocument(file);
    }

    async handleDrop(e) {
        e.preventDefault();
        this.documentViewport.classList.remove('drag-over');
        const file = e.dataTransfer.files[0];
        if (file && (file.name.endsWith('.docx') || file.name.endsWith('.doc'))) {
            await this.uploadDocument(file);
        } else {
            this.showToast('Please drop a Word document (.docx)', 'error');
        }
    }

    async uploadDocument(file) {
        try {
            const formData = new FormData();
            formData.append('file', file);
            const response = await fetch(`${this.apiBase}/upload`, { method: 'POST', body: formData });
            const data = await response.json();
            if (data.success) {
                this.docId = data.doc_id;
                this.content = data.content;
                this.styles = [];
                this.history = [];
                this.redoStack = [];
                this.savedSelection = null;
                this.renderDocument();
                this.updateStylesList();
                this.fileInfo.querySelector('.file-name').textContent = data.filename;
                this.showToast('Document uploaded successfully', 'success');
            } else {
                throw new Error(data.error);
            }
        } catch (error) {
            console.error('Upload error:', error);
            this.showToast('Failed to upload: ' + error.message, 'error');
        }
    }

    renderDocument() {
        if (!this.content.length) {
            this.emptyState.style.display = 'flex';
            this.documentContainer.style.display = 'none';
            return;
        }
        this.emptyState.style.display = 'none';
        this.documentContainer.style.display = 'block';
        this.documentContent.innerHTML = this.content.map((p, i) => `<p data-para="${i}">${this.escapeHtml(p.text)}</p>`).join('');
        this.applyAllStyles();
    }

    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    _iconHtml(s) {
        return s.svgCode
            ? `<span class="inline-icon" data-style-id="${s.id}" title="${this.escapeHtml(s.iconName)}">${s.svgCode}</span>`
            : `<img src="${s.iconData}" class="inline-icon" data-style-id="${s.id}" alt="">`;
    }

    _setupStepper(minusBtn, plusBtn, input, min, max, onChange) {
        const step = (dir) => {
            const val = Math.max(min, Math.min(max, parseInt(input.value) + dir));
            input.value = val;
            onChange(val);
        };
        const startHold = (dir) => {
            step(dir);
            let delay = setTimeout(() => {
                let interval = setInterval(() => step(dir), 60);
                const stop = () => { clearInterval(interval); document.removeEventListener('mouseup', stop); document.removeEventListener('mouseleave', stop); };
                document.addEventListener('mouseup', stop);
                document.addEventListener('mouseleave', stop);
                minusBtn._stopHold = plusBtn._stopHold = stop;
            }, 350);
            const cancelDelay = () => { clearTimeout(delay); if (minusBtn._stopHold) minusBtn._stopHold(); document.removeEventListener('mouseup', cancelDelay); };
            document.addEventListener('mouseup', cancelDelay);
        };
        minusBtn.addEventListener('mousedown', e => { e.preventDefault(); startHold(-1); });
        plusBtn.addEventListener('mousedown', e => { e.preventDefault(); startHold(1); });
    }

    resetStepperDefaults() {
        this.fontSize = '16px';
        this.letterSpacing = '4px';
        this.fontSizeInput.value = 16;
        this.letterSpacingInput.value = 4;
        this.ftFontSizeInput.value = 16;
    }

    toggleLetterSpacingPopover() {
        const isVisible = this.letterSpacingPopover.classList.contains('visible');
        if (isVisible) {
            this.letterSpacingPopover.classList.remove('visible');
            this.letterSpacingBtn.classList.remove('active');
        } else {
            this.letterSpacingPopover.classList.add('visible');
            this.letterSpacingBtn.classList.add('active');
        }
    }

    selectTool(tool) {
        if (!this.savedSelection) { this.showToast('Select text first', 'error'); return; }

        this.currentTool = null;
        this.toolButtons.forEach(btn => btn.classList.remove('active'));
        this.fontFamilyPopover.classList.remove('visible');
        this.fontFamilyBtn.classList.remove('active');
        this.fontOptions.forEach(o => o.classList.remove('active'));
        this.scriptSizePopover.classList.remove('visible');
        this.scriptSizeBtn.classList.remove('active');
        this.letterSpacingPopover.classList.remove('visible');
        this.letterSpacingBtn.classList.remove('active');
        this.hideFloatingToolbar();
        this.applyToolToSelection(tool);
    }

    toggleFontFamilyPopover() {
        const isVisible = this.fontFamilyPopover.classList.contains('visible');
        if (!isVisible && !this.savedSelection) { this.showToast('Select text first', 'error'); return; }
        this.currentTool = null;
        this.toolButtons.forEach(btn => btn.classList.remove('active'));
        this.scriptSizePopover.classList.remove('visible');
        this.scriptSizeBtn.classList.remove('active');
        this.letterSpacingPopover.classList.remove('visible');
        this.letterSpacingBtn.classList.remove('active');
        if (isVisible) {
            this.fontFamilyPopover.classList.remove('visible');
            this.fontFamilyBtn.classList.remove('active');
            this.fontOptions.forEach(o => o.classList.remove('active'));
        } else {
            this.fontFamilyPopover.classList.add('visible');
            this.fontFamilyBtn.classList.add('active');
        }
    }

    selectFontFamily(font) {
        this.fontOptions.forEach(o => o.classList.toggle('active', o.dataset.font === font));
        if (this.savedSelection) this.applyToolToSelection(font);
        this.fontFamilyPopover.classList.remove('visible');
        this.fontFamilyBtn.classList.remove('active');
        this.fontOptions.forEach(o => o.classList.remove('active'));
    }

    toggleScriptSizePopover() {
        const isVisible = this.scriptSizePopover.classList.contains('visible');
        if (!isVisible && !this.savedSelection) { this.showToast('Select text first', 'error'); return; }
        this.currentTool = null;
        this.toolButtons.forEach(btn => btn.classList.remove('active'));
        this.fontFamilyPopover.classList.remove('visible');
        this.fontFamilyBtn.classList.remove('active');
        this.letterSpacingPopover.classList.remove('visible');
        this.letterSpacingBtn.classList.remove('active');
        if (isVisible) {
            this.scriptSizePopover.classList.remove('visible');
            this.scriptSizeBtn.classList.remove('active');
        } else {
            this.scriptSizePopover.classList.add('visible');
            this.scriptSizeBtn.classList.add('active');
        }
    }

    selectScriptSize(type) {
        if (this.savedSelection) this.applyToolToSelection(type);
        this.scriptSizePopover.classList.remove('visible');
        this.scriptSizeBtn.classList.remove('active');
    }

    applyToolToSelection(tool, keepSelection = false) {
        if (!this.savedSelection) return;
        const { paraIndex, startOffset, endOffset, text } = this.savedSelection;

        const isDropcap = tool === 'dropcap';
        const sOff = isDropcap ? 0 : startOffset;
        const eOff = isDropcap ? 1 : endOffset;
        const sText = isDropcap ? text.charAt(0) : text;

        if (keepSelection) {
            const existing = this.styles.find(s => s.type === tool && s.paraIndex === paraIndex && s.startOffset === sOff && s.endOffset === eOff);
            if (existing) {
                existing.color = this.getColorForTool(tool);
                this.applyAllStyles();
                this.updateStylesList();
                this.restoreSelection(paraIndex, sOff, eOff);
                return;
            }
        }

        const style = {
            id: 'style-' + Date.now() + '-' + Math.random().toString(36).substr(2, 9),
            type: tool,
            text: sText,
            color: this.getColorForTool(tool),
            paraIndex,
            startOffset: sOff,
            endOffset: eOff,
            created_at: new Date().toISOString()
        };

        if (tool === 'callout') {
            style.bgColor = this.bgColor;
        }

        this.history.push({ action: 'add', style });
        this.redoStack = [];
        this.undoBtn.disabled = false;
        this.redoBtn.disabled = true;
        this.styles.push(style);
        this.logAction('add', style);

        const colorCat = this.getColorCategory(tool);
        if (colorCat) this.addColorToHistory(colorCat, style.color);
        if (tool === 'callout') this.addColorToHistory('bg', this.bgColor);

        this.applyAllStyles();
        this.updateStylesList();

        if (keepSelection) {
            this.restoreSelection(paraIndex, sOff, eOff);
        } else {
            this.savedSelection = null;
            this.resetStepperDefaults();
            window.getSelection().removeAllRanges();
            this.selectionHint.textContent = 'Select text to apply styles';
        }
    }

    restoreSelection(paraIndex, startOffset, endOffset) {
        const para = this.documentContent.querySelector(`p[data-para="${paraIndex}"]`);
        if (!para) return;
        const walker = document.createTreeWalker(para, NodeFilter.SHOW_TEXT, {
            acceptNode: n => n.parentElement?.closest('.inline-icon') ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT
        });
        let startNode = null, startNodeOffset = 0, endNode = null, endNodeOffset = 0;
        let current, total = 0;
        while ((current = walker.nextNode())) {
            const len = current.textContent.length;
            if (!startNode && total + len >= startOffset) { startNode = current; startNodeOffset = startOffset - total; }
            if (!endNode && total + len >= endOffset) { endNode = current; endNodeOffset = endOffset - total; break; }
            total += len;
        }
        if (startNode && endNode) {
            const range = document.createRange();
            range.setStart(startNode, startNodeOffset);
            range.setEnd(endNode, endNodeOffset);
            const sel = window.getSelection();
            sel.removeAllRanges();
            sel.addRange(range);
        }
    }

    openIconModal() {
        this.iconDescription.value = '';
        this.iconModal.style.display = 'flex';
        this.iconDescription.focus();
        this.setIconModalLoading(false);
    }

    closeIconModal() {
        this.iconModal.style.display = 'none';
        this.iconDescription.value = '';
    }

    setIconModalLoading(loading) {
        this.iconDescription.disabled = loading;
        this.iconModalSubmit.disabled = loading;
        this.iconModalSubmitText.textContent = loading ? 'Generating...' : 'Generate';
        this.iconModalSpinner.style.display = loading ? 'inline-block' : 'none';
    }

    async generateIcon() {
        const description = this.iconDescription.value.trim();
        if (!description) { this.showToast('Please describe the icon', 'error'); return; }

        this.setIconModalLoading(true);

        try {
            const response = await fetch(`${this.apiBase}/generate-icon`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ description })
            });
            const data = await response.json();
            if (!data.success) throw new Error(data.error || 'Failed to generate icon');

            this.closeIconModal();
            this.pendingIconData = { iconData: data.iconData, iconName: data.iconName, svgCode: data.svgCode };
            this.enterIconPlacementMode();
        } catch (error) {
            console.error('Icon generation error:', error);
            this.showToast('Failed to generate icon: ' + error.message, 'error');
            this.setIconModalLoading(false);
        }
    }

    enterIconPlacementMode() {
        this.documentContent.classList.add('icon-placement-mode');
        this.selectionHint.textContent = 'Click anywhere in text to place icon';
        this.showToast('Click in the document to place the icon', 'success');
    }

    exitIconPlacementMode() {
        this.pendingIconData = null;
        this.documentContent.classList.remove('icon-placement-mode');
        this.selectionHint.textContent = 'Select text to apply styles';
    }

    handleIconPlacement(e) {
        if (!this.pendingIconData) return;
        e.preventDefault();
        e.stopPropagation();

        const range = document.caretRangeFromPoint(e.clientX, e.clientY);
        if (!range) { this.showToast('Could not determine position', 'error'); return; }

        const para = this.getParentParagraph(range.startContainer);
        if (!para) { this.showToast('Click inside the text', 'error'); return; }

        const paraIndex = parseInt(para.dataset.para);
        const offset = this.getTextOffset(para, range.startContainer, range.startOffset);

        const style = {
            id: 'style-' + Date.now() + '-' + Math.random().toString(36).substr(2, 9),
            type: 'inlineicon',
            text: this.pendingIconData.iconName,
            color: '#000000',
            paraIndex,
            startOffset: offset,
            endOffset: offset,
            ...this.pendingIconData,
            created_at: new Date().toISOString()
        };

        this.history.push({ action: 'add', style });
        this.redoStack = [];
        this.undoBtn.disabled = false;
        this.redoBtn.disabled = true;
        this.styles.push(style);
        this.logAction('add', style);
        this.applyAllStyles();
        this.updateStylesList();
        this.showToast(`Icon "${style.iconName}" placed`, 'success');

        this.exitIconPlacementMode();
        window.getSelection().removeAllRanges();
    }

    handleTextSelection(e) {
        if (!this.docId) return;
        if (this.pendingIconData) return;
        if (e.target.closest('.toolbar') || e.target.closest('.styles-panel') || e.target.closest('.app-header') || e.target.closest('.modal-overlay') || e.target.closest('.floating-toolbar')) return;

        const selection = window.getSelection();
        if (!selection.rangeCount || selection.isCollapsed) {
            this.savedSelection = null;
            this.resetStepperDefaults();
            this.hideFloatingToolbar();
            this.selectionHint.textContent = 'Select text to apply styles';
            return;
        }

        const range = selection.getRangeAt(0);
        const selectedText = selection.toString().trim();
        if (!selectedText || !this.documentContent.contains(range.commonAncestorContainer)) {
            this.savedSelection = null;
            this.resetStepperDefaults();
            this.hideFloatingToolbar();
            this.selectionHint.textContent = 'Select text to apply styles';
            return;
        }

        const startNode = this.getParentParagraph(range.startContainer);
        if (!startNode) return;

        const paraIndex = parseInt(startNode.dataset.para);
        const startOffset = this.getTextOffset(startNode, range.startContainer, range.startOffset);
        const endOffset = this.getTextOffset(startNode, range.endContainer, range.endOffset);

        this.savedSelection = { paraIndex, startOffset, endOffset, text: selectedText };
        this.selectionHint.textContent = `"${selectedText.substring(0, 20)}${selectedText.length > 20 ? '...' : ''}" selected`;

        this.showFloatingToolbar(range);
    }

    showFloatingToolbar(range) {
        const rect = range.getBoundingClientRect();
        const toolbar = this.floatingToolbar;
        // Sync floating toolbar colors with current state
        const ftHighlight = toolbar.querySelector('.ft-highlight-icon');
        const ftTextcolor = toolbar.querySelector('.ft-textcolor-icon');
        if (ftHighlight) ftHighlight.style.background = this.bgColor;
        if (ftTextcolor) ftTextcolor.style.color = this.textColor;
        // Sync font size
        this.ftFontSizeInput.value = parseInt(this.fontSize);

        toolbar.classList.add('visible');
        toolbar.classList.remove('above');

        const tbRect = toolbar.getBoundingClientRect();
        const gap = 10;
        let left = rect.left + rect.width / 2 - tbRect.width / 2;
        let top = rect.top - tbRect.height - gap;
        let showAbove = true;

        // If not enough space above, show below
        if (top < 4) {
            top = rect.bottom + gap;
            showAbove = false;
        }

        // Clamp horizontal
        left = Math.max(4, Math.min(left, window.innerWidth - tbRect.width - 4));

        toolbar.style.left = left + 'px';
        toolbar.style.top = top + 'px';
        toolbar.classList.toggle('above', !showAbove);
    }

    hideFloatingToolbar() {
        this.floatingToolbar.classList.remove('visible');
        this.closeFtPopovers();
    }

    toggleFtPopover(popoverId) {
        const popover = this[popoverId];
        const isOpen = popover.classList.contains('visible');
        this.closeFtPopovers();
        if (!isOpen) popover.classList.add('visible');
    }

    closeFtPopovers() {
        this.ftFontFamilyPopover.classList.remove('visible');
        this.ftScriptSizePopover.classList.remove('visible');
    }

    getParentParagraph(node) {
        while (node && node !== this.documentContent) {
            if (node.nodeName === 'P' && node.dataset?.para !== undefined) return node;
            node = node.parentNode;
        }
        return null;
    }

    getTextOffset(paragraph, node, offset) {
        const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT, {
            acceptNode: n => n.parentElement?.closest('.inline-icon') ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT
        });
        let total = 0, current;
        while ((current = walker.nextNode())) {
            if (current === node) return total + offset;
            total += current.textContent.length;
        }
        return total + offset;
    }

    getColorForTool(tool) {
        if (tool === 'textcolor' || tool === 'dropcap') return this.textColor;
        if (tool === 'highlight') return this.bgColor;
        if (tool === 'fontsize') return this.fontSize;
        if (tool === 'letterspacing') return this.letterSpacing;
        return this.borderColor;
    }

    applyAllStyles() {
        this.documentContent.innerHTML = this.content.map((p, i) => `<p data-para="${i}">${this.escapeHtml(p.text)}</p>`).join('');

        const stylesByPara = {};
        this.styles.forEach(s => (stylesByPara[s.paraIndex] ??= []).push(s));

        for (const paraIndex in stylesByPara) {
            const paraStyles = stylesByPara[paraIndex];
            const para = this.documentContent.querySelector(`p[data-para="${paraIndex}"]`);
            if (!para) continue;

            const text = para.textContent;
            const iconStyles = paraStyles.filter(s => s.type === 'inlineicon');
            const calloutStyles = paraStyles.filter(s => s.type === 'callout');
            const textStyles = paraStyles.filter(s => s.type !== 'inlineicon' && s.type !== 'callout');

            const bounds = [...new Set([0, text.length, ...textStyles.flatMap(s => [Math.max(0, Math.min(s.startOffset, text.length)), Math.max(0, Math.min(s.endOffset, text.length))]), ...iconStyles.map(s => Math.max(0, Math.min(s.startOffset, text.length)))])].sort((a, b) => a - b);

            let result = '';
            for (let i = 0; i < bounds.length - 1; i++) {
                const [start, end] = [bounds[i], bounds[i + 1]];
                iconStyles.filter(s => s.startOffset === start).forEach(s => result += this._iconHtml(s));
                const seg = text.substring(start, end);
                if (!seg) continue;
                const active = textStyles.filter(s => s.startOffset <= start && s.endOffset >= end);
                result += active.length ? this.buildStyledSpan(active, seg) : this.escapeHtml(seg);
            }
            iconStyles.filter(s => s.startOffset >= text.length).forEach(s => result += this._iconHtml(s));
            para.innerHTML = result;

            if (calloutStyles.length) {
                const cs = calloutStyles[calloutStyles.length - 1];
                para.classList.add('callout-block');
                para.style.borderColor = cs.color;
                if (cs.bgColor) para.style.backgroundColor = cs.bgColor;
            }
        }
    }

    buildStyledSpan(styles, text) {
        const classes = ['styled-text', ...styles.map(s => s.type)];
        const inlineMap = { highlight: 'background-color', textcolor: 'color', border: 'border-color', circle: 'border-color', underline: 'text-decoration-color', wavyunderline: 'text-decoration-color', strikethrough: 'text-decoration-color', overline: 'text-decoration-color', fontsize: 'font-size', letterspacing: 'letter-spacing', dropcap: 'color' };
        const inline = styles.map(s => inlineMap[s.type] ? `${inlineMap[s.type]}:${s.color}` : null).filter(Boolean);
        const ids = styles.map(s => s.id).join(',');
        const styleAttr = inline.length ? ` style="${inline.join(';')}"` : '';
        return `<span class="${classes.join(' ')}" data-style-id="${ids}"${styleAttr}>${this.escapeHtml(text)}</span>`;
    }

    updateStylesList() {
        const count = this.styles.length;
        this.styleCount.textContent = count;
        if (!count) {
            this.stylesList.innerHTML = '<div class="empty-styles"><p>No styles applied</p><small>Select text and apply styles</small></div>';
            return;
        }
        const icons = { bold: '<strong>B</strong>', italic: '<em>I</em>', underline: '<u>U</u>', wavyunderline: '<span style="text-decoration:underline wavy">W</span>', strikethrough: '<s>S</s>', superscript: 'X\u00B2', subscript: 'X\u2082', highlight: '▮', textcolor: 'A', border: '□', circle: '○', sansserif: 'Aa', mono: 'T_', rounded: 'Rr', smallcaps: 'Aᴀ', fontsize: 'Tt', inlineicon: '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><path d="M21 15l-5-5L5 21"/></svg>', letterspacing: '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><text x="1" y="14" font-size="12" fill="currentColor" stroke="none">A</text><text x="15" y="14" font-size="12" fill="currentColor" stroke="none">V</text><line x1="2" y1="20" x2="22" y2="20"/><polyline points="5 22 2 20 5 18"/><polyline points="19 22 22 20 19 18"/></svg>', overline: 'O̅', callout: '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="3" width="20" height="18" rx="3"/><line x1="6" y1="8" x2="18" y2="8"/><line x1="6" y1="12" x2="18" y2="12"/><line x1="6" y1="16" x2="14" y2="16"/></svg>', dropcap: '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><text x="1" y="17" font-size="20" font-weight="bold" fill="currentColor" stroke="none">A</text><line x1="15" y1="5" x2="23" y2="5"/><line x1="15" y1="10" x2="23" y2="10"/><line x1="15" y1="15" x2="23" y2="15"/><line x1="1" y1="22" x2="23" y2="22"/></svg>' };
        this.stylesList.innerHTML = this.styles.map(s => {
            const noColorIcon = ['fontsize', 'inlineicon', 'letterspacing'];
            const iconStyle = noColorIcon.includes(s.type) ? '' : ` style="color:${s.color}"`;
            const displayType = s.type === 'fontsize' ? `font size (${s.color})` : s.type === 'inlineicon' ? 'inline icon' : s.type === 'letterspacing' ? `letter spacing (${s.color})` : s.type === 'callout' ? 'callout' : s.type === 'dropcap' ? 'drop cap' : s.type === 'wavyunderline' ? 'wavy underline' : s.type;
            return `
            <div class="style-item" data-id="${s.id}">
                <div class="style-icon"${iconStyle}>${icons[s.type] || '•'}</div>
                <div class="style-details">
                    <div class="style-type">${displayType}</div>
                    <div class="style-preview">"${s.text.substring(0, 20)}${s.text.length > 20 ? '...' : ''}"</div>
                </div>
                <button class="style-delete" title="Delete">
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
                </button>
            </div>
        `}).join('');
        this.stylesList.querySelectorAll('.style-delete').forEach(btn => {
            btn.addEventListener('click', e => { e.stopPropagation(); this.deleteStyle(e.currentTarget.closest('.style-item').dataset.id); });
        });
    }

    deleteStyle(id) {
        const index = this.styles.findIndex(s => s.id === id);
        if (index !== -1) {
            const style = this.styles[index];
            this.history.push({ action: 'delete', style });
            this.redoStack = [];
            this.undoBtn.disabled = false;
            this.redoBtn.disabled = true;
            this.styles.splice(index, 1);
            this.logAction('delete', style);
            this.applyAllStyles();
            this.updateStylesList();
        }
    }

    undo() {
        if (!this.history.length) return;
        const last = this.history.pop();
        if (last.action === 'add') {
            const idx = this.styles.findIndex(s => s.id === last.style.id);
            if (idx !== -1) this.styles.splice(idx, 1);
        } else if (last.action === 'delete') {
            this.styles.push(last.style);
        } else if (last.action === 'clear') {
            this.styles = last.styles;
        }
        this.redoStack.push(last);
        this.applyAllStyles();
        this.updateStylesList();
        this.undoBtn.disabled = !this.history.length;
        this.redoBtn.disabled = false;
    }

    redo() {
        if (!this.redoStack.length) return;
        const last = this.redoStack.pop();
        if (last.action === 'add') {
            this.styles.push(last.style);
        } else if (last.action === 'delete') {
            const idx = this.styles.findIndex(s => s.id === last.style.id);
            if (idx !== -1) this.styles.splice(idx, 1);
        } else if (last.action === 'clear') {
            this.styles = [];
        }
        this.history.push(last);
        this.applyAllStyles();
        this.updateStylesList();
        this.undoBtn.disabled = false;
        this.redoBtn.disabled = !this.redoStack.length;
    }

    async clearAllStyles() {
        if (!this.styles.length) { this.showToast('No styles to clear', 'error'); return; }
        if (!confirm('Clear all styles?')) return;
        this.history.push({ action: 'clear', styles: [...this.styles] });
        this.redoStack = [];
        this.undoBtn.disabled = false;
        this.redoBtn.disabled = true;
        const count = this.styles.length;
        this.styles = [];
        this.logAction('clear', null, count);
        this.applyAllStyles();
        this.updateStylesList();
        this.showToast('All styles cleared', 'success');
    }

    async saveStyles() {
        if (!this.docId) { this.showToast('No document loaded', 'error'); return; }
        try {
            const res = await fetch(`${this.apiBase}/document/${this.docId}/styles`, {
                method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ styles: this.styles })
            });
            const data = await res.json();
            if (!data.success) throw new Error(data.error);

            const exportRes = await fetch(`${this.apiBase}/document/${this.docId}/export`, {
                method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ styles: this.styles })
            });
            const exportData = await exportRes.json();
            if (exportData.success) {
                this.showToast('Image exported! Downloading...', 'success');
                window.location.href = `${this.apiBase}/document/${this.docId}/download`;
            } else {
                throw new Error(exportData.error);
            }
        } catch (error) {
            console.error('Save error:', error);
            this.showToast('Failed to save: ' + error.message, 'error');
        }
    }

    async logAction(action, style = null, stylesCleared = null) {
        if (!this.docId) return;
        const entry = {
            log_id: 'log-' + Date.now() + '-' + Math.random().toString(36).substr(2, 9),
            action, timestamp: new Date().toISOString(),
            style: style ? { id: style.id, type: style.type, text: style.text, color: style.color, paraIndex: style.paraIndex, startOffset: style.startOffset, endOffset: style.endOffset, ...(style.bgColor ? { bgColor: style.bgColor } : {}) } : null,
            styles_cleared: stylesCleared
        };
        try {
            await fetch(`${this.apiBase}/document/${this.docId}/log`, {
                method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(entry)
            });
        } catch (e) { console.error('Failed to log action:', e); }
    }

    handleKeyboard(e) {
        if (e.target.tagName === 'INPUT' || e.target.tagName === 'SELECT') return;
        const mod = e.ctrlKey || e.metaKey;
        const shortcuts = { z: () => this.undo(), y: () => this.redo(), s: () => this.saveStyles(), b: () => this.selectTool('bold'), i: () => this.selectTool('italic'), u: () => this.selectTool('underline') };
        const keys = { h: 'highlight', t: 'textcolor', r: 'border', c: 'circle', o: 'overline', w: 'wavyunderline', q: 'callout', d: 'dropcap' };

        if (mod && shortcuts[e.key]) { e.preventDefault(); shortcuts[e.key](); }
        else if (!mod && keys[e.key?.toLowerCase()]) this.selectTool(keys[e.key.toLowerCase()]);
        else if (e.key === 'Escape') {
            if (this.pendingIconData) {
                this.exitIconPlacementMode();
                return;
            }
            if (this.iconModal.style.display !== 'none') {
                this.closeIconModal();
                return;
            }
            this.currentTool = null;
            this.savedSelection = null;
            this.resetStepperDefaults();
            this.toolButtons.forEach(btn => btn.classList.remove('active'));
            this.fontFamilyPopover.classList.remove('visible');
            this.fontFamilyBtn.classList.remove('active');
            this.scriptSizePopover.classList.remove('visible');
            this.scriptSizeBtn.classList.remove('active');
            this.letterSpacingPopover.classList.remove('visible');
            this.letterSpacingBtn.classList.remove('active');
            this.closeAllColorPopovers();
            this.hideFloatingToolbar();
            this.selectionHint.textContent = 'Select text to apply styles';
            window.getSelection().removeAllRanges();
        }
    }

    loadColorHistory(category) {
        try { return (JSON.parse(localStorage.getItem(this.colorHistoryKeys[category]) || '[]')).filter(c => typeof c === 'string' && /^#[0-9a-fA-F]{6}$/.test(c)).slice(0, this.COLOR_HISTORY_MAX); }
        catch (e) { return []; }
    }

    saveColorHistory(category) {
        try { localStorage.setItem(this.colorHistoryKeys[category], JSON.stringify(this.colorHistory[category])); } catch (e) { /* ignore */ }
    }

    addColorToHistory(category, color) {
        if (!color || !/^#[0-9a-fA-F]{6}$/i.test(color)) return;
        const normalized = color.toLowerCase();
        const history = this.colorHistory[category];
        const idx = history.indexOf(normalized);
        if (idx !== -1) history.splice(idx, 1);
        history.unshift(normalized);
        if (history.length > this.COLOR_HISTORY_MAX) history.pop();
        this.saveColorHistory(category);
        this.renderColorSwatches(category);
    }

    getColorCategory(tool) {
        if (tool === 'textcolor' || tool === 'dropcap') return 'text';
        if (tool === 'highlight') return 'bg';
        if (tool === 'fontsize' || tool === 'letterspacing') return null;
        return 'border';
    }

    renderColorSwatches(category) {
        const container = { text: this.swatchesText, bg: this.swatchesBg, border: this.swatchesBorder }[category];
        if (!container) return;
        container.innerHTML = '';
        this.colorHistory[category].forEach(color => {
            const swatch = document.createElement('div');
            swatch.className = 'color-swatch';
            swatch.style.backgroundColor = color;
            swatch.title = color;
            swatch.addEventListener('click', () => this.selectSwatchColor(category, color));
            container.appendChild(swatch);
        });
    }

    selectSwatchColor(category, color) {
        if (category === 'text') {
            this.textColor = color;
            this.pickerText.value = color;
            this.textcolorIcon.style.color = color;
            this.previewText.style.backgroundColor = color;
        } else if (category === 'bg') {
            this.bgColor = color;
            this.pickerBg.value = color;
            this.highlightIcon.style.background = color;
            this.previewBg.style.backgroundColor = color;
        } else {
            this.borderColor = color;
            this.pickerBorder.value = color;
            this.previewBorder.style.backgroundColor = color;
        }
        this.closeAllColorPopovers();
    }

    toggleColorPopover(picker) {
        const popover = this.colorPopovers[picker];
        const isOpen = popover.classList.contains('visible');
        this.closeAllColorPopovers();
        if (!isOpen) {
            popover.classList.add('visible');
            popover.closest('.color-picker-item').querySelector('.color-trigger').classList.add('active');
        }
    }

    closeAllColorPopovers() {
        Object.values(this.colorPopovers).forEach(p => p.classList.remove('visible'));
        this.colorTriggers.forEach(t => t.classList.remove('active'));
    }

    showToast(message, type = 'success') {
        const toast = document.createElement('div');
        toast.className = `toast ${type}`;
        const icon = type === 'success'
            ? '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>'
            : '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg>';
        toast.innerHTML = `<span class="toast-icon">${icon}</span><span>${message}</span>`;
        this.toastContainer.appendChild(toast);
        setTimeout(() => { toast.style.animation = 'slideIn 0.2s ease reverse forwards'; setTimeout(() => toast.remove(), 200); }, 3000);
    }
}

document.addEventListener('DOMContentLoaded', () => { window.docTypography = new DocumentTypography(); });
