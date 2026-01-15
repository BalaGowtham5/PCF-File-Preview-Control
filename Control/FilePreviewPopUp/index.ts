import { IInputs, IOutputs } from './generated/ManifestTypes';
import * as XLSX from 'xlsx';
import { renderAsync } from 'docx-preview'; 

export class FilePreviewPopUp implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container: HTMLDivElement;
    private _button: HTMLButtonElement;
    private _modal: HTMLDivElement;
    private _contentContainer: HTMLDivElement;
    private _headerTitle: HTMLSpanElement;
    
    private _logicalName = "";    
    private _fileUrl = "";
    private _displayFileName = "File"; 
    private _modalAttached = false; 

    constructor() {
        // Empty constructor required by PCF
    }

    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
        this._container = container;

        // 1. Create Button
        this._button = document.createElement("button");
        this._button.className = "ccwis-preview-btn";
        this._button.innerText = "Initializing...";
        this._button.disabled = true;
        this._button.addEventListener("click", this.openModal.bind(this));

        this._container.appendChild(this._button);

        // 2. Build Modal
        this.createModalElements();
    }

    private createModalElements(): void {
        this._modal = document.createElement("div");
        this._modal.className = "ccwis-modal-overlay";
        this._modal.style.display = "none";

        const content = document.createElement("div");
        content.className = "ccwis-modal-content";

        const header = document.createElement("div");
        header.className = "ccwis-modal-header";
        
        this._headerTitle = document.createElement("span");
        this._headerTitle.style.fontWeight = "bold";
        this._headerTitle.innerText = "Preview";

        const closeBtn = document.createElement("button");
        closeBtn.className = "ccwis-close-btn";
        closeBtn.innerText = "✖ Close";
        closeBtn.onclick = this.closeModal.bind(this);
        
        header.appendChild(this._headerTitle);
        header.appendChild(closeBtn);

        this._contentContainer = document.createElement("div");
        this._contentContainer.className = "ccwis-modal-frame";

        content.appendChild(header);
        content.appendChild(this._contentContainer);
        this._modal.appendChild(content);
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._logicalName = context.parameters.targetFileField.raw || "";
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const id = (context.mode as any).contextInfo.entityId;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const etn = (context.mode as any).contextInfo.entityTypeName;

        if (!id || id === '00000000-0000-0000-0000-000000000000') {
            this._button.innerText = "Save Record First";
            this._button.disabled = true;
            return;
        }

        // Setup URL
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const clientUrl = (context as any).page.getClientUrl();
        this._fileUrl = clientUrl + "/api/data/v9.2/" + etn + "s(" + id + ")/" + this._logicalName + "/$value";

        this._button.disabled = false;
        this._button.innerText = "📄 Preview File";

        // Optional Name Fetch
        context.webAPI.retrieveRecord(etn, id, "?$select=" + this._logicalName)
            .then(result => {
                let name = result[this._logicalName];
                if (!name || this.isGuid(name)) name = result[this._logicalName + "_name"];
                if (name && !this.isGuid(name)) {
                    this._displayFileName = name;
                    this._button.innerText = "📄 Preview " + name;
                }
                return null;
            }).catch(() => {
                // Ignore errors
            });
    }

    private isGuid(val: string): boolean {
        return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(val);
    }

    private openModal(): void {
        // Attach to BODY
        if (!this._modalAttached) {
            document.body.appendChild(this._modal);
            this._modalAttached = true;
        }
        
        this._modal.style.display = "flex";
        this._headerTitle.innerText = this._displayFileName;
        this._contentContainer.innerHTML = "<h3 style='margin-top:20px;'>Scanning...</h3>";

        fetch(this._fileUrl)
            .then(response => {
                // CUSTOM HANDLE: 404 means no file is attached
                if (response.status === 404) {
                    this.renderEmptyState();
                    return null; 
                }
                
                if (!response.ok) throw new Error("Server Error: " + response.status);
                
                const type = response.headers.get('content-type') || "";
                return response.arrayBuffer().then(buffer => ({ buffer, type }));
            })
            .then((data) => {
                if (!data) return null; // Stopped at 404
                
                const detectedType = this.detectType(data.buffer, data.type);
                this.renderContent(data.buffer, detectedType);
                return null;
            })
            .catch(err => {
                this._contentContainer.innerHTML = `<p style='color:red'>Error: ${err.message}</p>`;
            });
    }

    // NEW: Handles Empty Files
    private renderEmptyState(): void {
        this._contentContainer.innerHTML = `
            <div style="text-align:center; padding:50px;">
                <h3>Please attach a file.</h3>
                <p style="color:#666;">No document found in this field.</p>
            </div>`;
    }

    private detectType(buffer: ArrayBuffer, serverType: string): string {
        const bytes = new Uint8Array(buffer).subarray(0, 4);
        let header = "";
        for (const byte of bytes) {
            header += byte.toString(16).toUpperCase().padStart(2, '0');
        }

        // 1. Magic Bytes
        if (header.startsWith("25504446")) return "pdf"; 
        if (header.startsWith("FFD8FF")) return "image"; 
        if (header.startsWith("89504E47")) return "image"; 
        if (header.startsWith("47494638")) return "image"; 
        if (header.startsWith("504B0304")) return "zip"; 

        // 2. Extension Check
        const name = this._displayFileName.toLowerCase();
        if (name.endsWith(".txt")) return "text";
        if (name.endsWith(".csv")) return "excel"; 
        
        // 3. Server Header
        if (serverType.includes("text/plain")) return "text";
        if (serverType.includes("csv")) return "excel";

        // 4. Brute Force Text
        if (!this.isBinary(buffer)) return "text";

        return "unknown";
    }

    private isBinary(buffer: ArrayBuffer): boolean {
        const bytes = new Uint8Array(buffer).subarray(0, 1000); 
        for (const byte of bytes) {
            if (byte === 0) return true; 
        }
        return false;
    }

    private renderContent(buffer: ArrayBuffer, magicType: string): void {
        this._contentContainer.innerHTML = ""; 
        let ext = magicType;

        if (magicType === "zip") {
             if (this._displayFileName.toLowerCase().endsWith("xlsx")) ext = "excel";
             else if (this._displayFileName.toLowerCase().endsWith("docx")) ext = "word";
             else ext = "try_word_then_excel"; 
        }

        if (ext === 'pdf') this.renderPdf(buffer);
        else if (ext === 'image') this.renderImage(buffer);
        else if (ext === 'excel') this.renderExcel(buffer);
        else if (ext === 'word') this.renderWord(buffer);
        else if (ext === 'try_word_then_excel') this.renderWord(buffer, true); 
        else if (ext === 'text') this.renderText(buffer);
        else this.showUnsupportedMessage("Unknown Format");
    }

    private renderPdf(buffer: ArrayBuffer): void {
        const blob = new Blob([buffer], { type: 'application/pdf' });
        const url = URL.createObjectURL(blob);
        const iframe = document.createElement("iframe");
        iframe.src = url;
        iframe.style.width = "100%";
        iframe.style.height = "100%";
        iframe.style.border = "none";
        this._contentContainer.appendChild(iframe);
    }

    private renderImage(buffer: ArrayBuffer): void {
        const blob = new Blob([buffer]);
        const url = URL.createObjectURL(blob);
        const img = document.createElement("img");
        img.src = url;
        img.style.maxWidth = "100%";
        img.style.height = "auto";
        img.style.marginTop = "10px";
        this._contentContainer.appendChild(img);
    }

    private renderWord(buffer: ArrayBuffer, retryAsExcel = false): void {
        const options = {
            className: "docx_viewer", 
            inWrapper: true, 
            ignoreWidth: false,
            ignoreHeight: false,
            ignoreFonts: false, 
            breakPages: true,
            ignoreLastRenderedPageBreak: false,
            experimental: false,
            trimXmlDeclaration: true,
            debug: false,
        };
        
        renderAsync(buffer, this._contentContainer, undefined, options)
            .then(() => { 
                console.log("Word Rendered");
                return null;
            })
            .catch((err) => {
                if(retryAsExcel) this.renderExcel(buffer);
                else this.showUnsupportedMessage("Word Render Failed");
            });
    }

    private renderExcel(buffer: ArrayBuffer): void {
        try {
            const workbook = XLSX.read(buffer, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const html = XLSX.utils.sheet_to_html(firstSheet);
            
            const excelStyle = `
                <style>
                    .excel-container { 
                        width: 100%; 
                        background: white; 
                        padding: 10px; 
                        overflow: auto; 
                        display: block;
                        max-height: 100%;
                    }
                    .excel-container table { border-collapse: collapse; font-family: Calibri, sans-serif; font-size: 14px; }
                    .excel-container td, .excel-container th { border: 1px solid #ccc; padding: 4px 8px; white-space: nowrap; }
                    .excel-container th { background-color: #f3f3f3; font-weight: bold; text-align: center; }
                </style>
            `;

            this._contentContainer.innerHTML = excelStyle + `<div class="excel-container">${html}</div>`;
        } catch (e) {
            this.showUnsupportedMessage("Excel Parse Failed");
        }
    }

    private renderText(buffer: ArrayBuffer): void {
        try {
            const decoder = new TextDecoder("utf-8");
            const text = decoder.decode(buffer);
            this._contentContainer.innerHTML = `
                <div style="padding:40px; background:white; width:90%; height:100%; overflow:auto; text-align:left; box-shadow: 0 0 10px rgba(0,0,0,0.1);">
                    <pre style="white-space: pre-wrap; font-family: 'Consolas', 'Courier New', monospace; font-size: 14px; color: #333;">${text}</pre>
                </div>`;
        } catch (e) {
            this.showUnsupportedMessage("Text Decode Failed");
        }
    }

    // NEW: Handles Unsupported Files (No Download Button)
    private showUnsupportedMessage(reason: string): void {
        this._contentContainer.innerHTML = `
            <div style="text-align:center; padding:50px;">
                <h3>${reason}</h3>
                <p style="color:#444; margin-top:15px; font-size:16px;">
                    File format is not supported for preview. Please close and download the file.
                </p>
            </div>`;
    }

    private closeModal(): void {
        this._modal.style.display = "none";
        this._contentContainer.innerHTML = "";
        
        if (this._modalAttached && document.body.contains(this._modal)) {
            document.body.removeChild(this._modal);
            this._modalAttached = false;
        }
    }

    public getOutputs(): IOutputs { return {}; }
    
    public destroy(): void {
        if (this._modalAttached && document.body.contains(this._modal)) {
            document.body.removeChild(this._modal);
        }
    }
}