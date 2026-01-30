class DocumentTypography {
    constructor() {
        this.docId = null;
        this.content = [];
        this.styles = [];
        this.history = [];
        this.currentTool = null;
        this.highlightColor = '#FFEB3B';
        this.textColor = '#E53935';
        this.borderColor = '#1976D2';
        this.apiBase = 'http://localhost:5001/api';
        this.initElements();
        this.initEventListeners();
    }

    initElements() {
        this.fileInput = document.getElementById('fileInput');
        this.uploadBtn = document.getElementById('uploadBtn');
        this.uploadBtnAlt = document.getElementById('uploadBtnAlt');
        this.documentViewport = document.getElementById('documentViewport');
        this.documentContainer = document.getElementById('documentContainer');
        this.documentContent = document.getElementById('documentContent');
        this.emptyState = document.getElementById('emptyState');
        this.fileInfo = document.getElementById('fileInfo');
        this.selectionHint = document.getElementById('selectionHint');
        this.toolButtons = document.querySelectorAll('.tool-btn');
        this.highlightColorInput = document.getElementById('highlightColor');
        this.textColorInput = document.getElementById('textColor');
        this.borderColorInput = document.getElementById('borderColor');
        this.quickColorBtns = document.querySelectorAll('.quick-color');
        this.undoBtn = document.getElementById('undoBtn');
        this.clearBtn = document.getElementById('clearBtn');
        this.saveBtn = document.getElementById('saveBtn');
        this.stylesList = document.getElementById('stylesList');
        this.styleCount = document.getElementById('styleCount');
        this.toastContainer = document.getElementById('toastContainer');
    }

    initEventListeners() {
        this.uploadBtn.addEventListener('click', () => this.fileInput.click());
        this.uploadBtnAlt.addEventListener('click', () => this.fileInput.click());
        this.fileInput.addEventListener('change', e => this.handleFileUpload(e));
        this.documentViewport.addEventListener('dragover', e => { e.preventDefault(); this.documentViewport.classList.add('drag-over'); });
        this.documentViewport.addEventListener('dragleave', e => { e.preventDefault(); this.documentViewport.classList.remove('drag-over'); });
        this.documentViewport.addEventListener('drop', e => this.handleDrop(e));
        this.toolButtons.forEach(btn => btn.addEventListener('click', () => this.selectTool(btn.dataset.tool)));
        this.highlightColorInput.addEventListener('input', e => this.highlightColor = e.target.value);
        this.textColorInput.addEventListener('input', e => this.textColor = e.target.value);
        this.borderColorInput.addEventListener('input', e => this.borderColor = e.target.value);
        this.quickColorBtns.forEach(btn => btn.addEventListener('click', () => {
            this.highlightColor = btn.dataset.color;
            this.highlightColorInput.value = btn.dataset.color;
        }));
        this.undoBtn.addEventListener('click', () => this.undo());
        this.clearBtn.addEventListener('click', () => this.clearAllStyles());
        this.saveBtn.addEventListener('click', () => this.saveStyles());
        document.addEventListener('mouseup', e => this.handleTextSelection(e));
        document.addEventListener('keydown', e => this.handleKeyboard(e));
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

    selectTool(tool) {
        if (this.currentTool === tool) {
            this.currentTool = null;
            this.toolButtons.forEach(btn => btn.classList.remove('active'));
            this.selectionHint.textContent = 'Select text to apply styles';
        } else {
            this.currentTool = tool;
            this.toolButtons.forEach(btn => btn.classList.toggle('active', btn.dataset.tool === tool));
            this.selectionHint.textContent = `Select text to apply ${tool}`;
        }
    }

    handleTextSelection() {
        if (!this.currentTool || !this.docId) return;
        const selection = window.getSelection();
        if (!selection.rangeCount || selection.isCollapsed) return;
        const range = selection.getRangeAt(0);
        const selectedText = selection.toString().trim();
        if (!selectedText || !this.documentContent.contains(range.commonAncestorContainer)) return;

        const startNode = this.getParentParagraph(range.startContainer);
        const endNode = this.getParentParagraph(range.endContainer);
        if (!startNode || !endNode) return;

        const style = {
            id: 'style-' + Date.now() + '-' + Math.random().toString(36).substr(2, 9),
            type: this.currentTool,
            text: selectedText,
            color: this.getColorForTool(this.currentTool),
            paraIndex: parseInt(startNode.dataset.para),
            startOffset: this.getTextOffset(startNode, range.startContainer, range.startOffset),
            endOffset: this.getTextOffset(startNode, range.endContainer, range.endOffset),
            created_at: new Date().toISOString()
        };

        this.history.push({ action: 'add', style });
        this.undoBtn.disabled = false;
        this.styles.push(style);
        this.logAction('add', style);
        this.applyAllStyles();
        this.updateStylesList();
        selection.removeAllRanges();
    }

    getParentParagraph(node) {
        while (node && node !== this.documentContent) {
            if (node.nodeName === 'P' && node.dataset?.para !== undefined) return node;
            node = node.parentNode;
        }
        return null;
    }

    getTextOffset(paragraph, node, offset) {
        const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT, null, false);
        let total = 0, current;
        while ((current = walker.nextNode())) {
            if (current === node) return total + offset;
            total += current.textContent.length;
        }
        return total + offset;
    }

    getColorForTool(tool) {
        if (tool === 'highlight') return this.highlightColor;
        if (tool === 'textcolor') return this.textColor;
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
            const bounds = [...new Set([0, text.length, ...paraStyles.flatMap(s => [Math.max(0, Math.min(s.startOffset, text.length)), Math.max(0, Math.min(s.endOffset, text.length))])])].sort((a, b) => a - b);

            let result = '';
            for (let i = 0; i < bounds.length - 1; i++) {
                const [start, end] = [bounds[i], bounds[i + 1]];
                const seg = text.substring(start, end);
                if (!seg) continue;
                const active = paraStyles.filter(s => s.startOffset <= start && s.endOffset >= end);
                result += active.length ? this.buildStyledSpan(active, seg) : this.escapeHtml(seg);
            }
            para.innerHTML = result;
        }
    }

    buildStyledSpan(styles, text) {
        const classes = ['styled-text', ...styles.map(s => s.type)];
        const inlineMap = { highlight: 'background-color', textcolor: 'color', border: 'border-color', circle: 'border-color', underline: 'text-decoration-color', strikethrough: 'text-decoration-color' };
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
        const icons = { bold: '<strong>B</strong>', italic: '<em>I</em>', underline: '<u>U</u>', strikethrough: '<s>S</s>', highlight: '▮', textcolor: 'A', border: '□', circle: '○' };
        this.stylesList.innerHTML = this.styles.map(s => `
            <div class="style-item" data-id="${s.id}">
                <div class="style-icon" style="color:${s.color}">${icons[s.type] || '•'}</div>
                <div class="style-details">
                    <div class="style-type">${s.type}</div>
                    <div class="style-preview">"${s.text.substring(0, 20)}${s.text.length > 20 ? '...' : ''}"</div>
                </div>
                <button class="style-delete" title="Delete">
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
                </button>
            </div>
        `).join('');
        this.stylesList.querySelectorAll('.style-delete').forEach(btn => {
            btn.addEventListener('click', e => { e.stopPropagation(); this.deleteStyle(e.currentTarget.closest('.style-item').dataset.id); });
        });
    }

    deleteStyle(id) {
        const index = this.styles.findIndex(s => s.id === id);
        if (index !== -1) {
            const style = this.styles[index];
            this.history.push({ action: 'delete', style });
            this.undoBtn.disabled = false;
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
        this.applyAllStyles();
        this.updateStylesList();
        this.undoBtn.disabled = !this.history.length;
    }

    async clearAllStyles() {
        if (!this.styles.length) { this.showToast('No styles to clear', 'error'); return; }
        if (!confirm('Clear all styles?')) return;
        this.history.push({ action: 'clear', styles: [...this.styles] });
        this.undoBtn.disabled = false;
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
            style: style ? { id: style.id, type: style.type, text: style.text, color: style.color, paraIndex: style.paraIndex, startOffset: style.startOffset, endOffset: style.endOffset } : null,
            styles_cleared: stylesCleared
        };
        try {
            await fetch(`${this.apiBase}/document/${this.docId}/log`, {
                method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(entry)
            });
        } catch (e) { console.error('Failed to log action:', e); }
    }

    handleKeyboard(e) {
        if (e.target.tagName === 'INPUT') return;
        const mod = e.ctrlKey || e.metaKey;
        const shortcuts = { z: () => this.undo(), s: () => this.saveStyles(), b: () => this.selectTool('bold'), i: () => this.selectTool('italic'), u: () => this.selectTool('underline') };
        const keys = { h: 'highlight', t: 'textcolor', r: 'border', c: 'circle' };

        if (mod && shortcuts[e.key]) { e.preventDefault(); shortcuts[e.key](); }
        else if (!mod && keys[e.key?.toLowerCase()]) this.selectTool(keys[e.key.toLowerCase()]);
        else if (e.key === 'Escape') {
            this.currentTool = null;
            this.toolButtons.forEach(btn => btn.classList.remove('active'));
            this.selectionHint.textContent = 'Select text to apply styles';
        }
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
