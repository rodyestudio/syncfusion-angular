import { Component, ViewEncapsulation, ViewChild } from '@angular/core';
import { ToolbarService, DocumentEditorContainerComponent } from '@syncfusion/ej2-angular-documenteditor';
import { TitleBar } from '../title-bar';
import { defaultDocument, WEB_API_ACTION } from '../data';
import { isNullOrUndefined } from '@syncfusion/ej2-base';
import {convertToHtml} from "mammoth/mammoth.browser";
import { defaultKopSurat, defaultTtdSurat, defaultKopSuratTanpaLogo } from '../kopsurat';

/**
 * Document Editor Component
 */
@Component({
    selector: 'app-root',
    templateUrl: 'app.component.html',
    encapsulation: ViewEncapsulation.None,
    providers: [ToolbarService]
})
export class AppComponent {
    public hostUrl: string = 'https://ej2services.syncfusion.com/production/web-services/';
    @ViewChild('documenteditor_default')
    public container: DocumentEditorContainerComponent;
    titleBar: TitleBar;
    @ViewChild('document_editor')
    public documentViewer: DocumentEditorContainerComponent;

    onCreate(): void {
        //let titleBarElement: HTMLElement = document.getElementById('default_title_bar');
        //this.titleBar = new TitleBar(titleBarElement, this.container.documentEditor, true);
        // this.container.documentEditor.open(JSON.stringify(defaultDocument));
        //this.titleBar.updateDocumentTitle();    

        this.container.locale = 'en-US';
        this.container.serviceUrl = this.hostUrl + WEB_API_ACTION;
        this.container.documentEditor.documentName = 'Getting Started';
        this.container.documentEditor.focusIn();        
        //this.container.toolbarModule.toolbar.hideItem(14, true);    //note: custom toolbar pagesetup
        //this.container.documentEditor.selection.sectionFormat.pageWidth = 595.3;  //note: default paper page A4
        //this.container.documentEditor.selection.sectionFormat.pageHeight = 841.9; // note: default paper page A4

        // let styleJson: any = {
        //     "type": "Character",
        //     "name": "New CharacterStyle",
        //     "basedOn": "Default Paragraph Font",
        //     "characterFormat": {
        //         "fontSize": 16.0,
        //         "fontFamily": "Calibri Light",
        //         "fontColor": "#2F5496",
        //         "bold": true,
        //         "italic": true,
        //         "underline": "Single"
        //     }
        // };
        
        // this.container.documentEditor.editor.createStyle(JSON.stringify(styleJson));
        // this.container.documentEditor.characterFormat.fontFamily = "Calibri Light";
        // this.container.documentEditor.characterFormat.fontSize = 16.0;
    }

    onCreatePreview(): void {
        let titleBarElement: HTMLElement = document.getElementById('default_title_bar');
        this.titleBar = new TitleBar(titleBarElement, this.documentViewer.documentEditor, true);
        this.titleBar.updateDocumentTitle();   
    }

    onDocumentChange(): void {
        if (!isNullOrUndefined(this.titleBar)) {
            this.titleBar.updateDocumentTitle();
        }
        this.container.documentEditor.focusIn();
    }

    onSelectionChange(): void {
        //this.container.documentEditor.selection.paragraphFormat.firstLineIndent= 12;
        //this.container.documentEditor.selection.paragraphFormat.afterSpacing = 24;
        //this.container.documentEditor.selection.paragraphFormat.beforeSpacing = 24;
        //this.container.documentEditor.selection.paragraphFormat.lineSpacingType='AtLeast';
        //this.container.documentEditor.selection.paragraphFormat.lineSpacing= 6;
        //this.container.documentEditor.selection.paragraphFormat.leftIndent= 8;
        //this.container.documentEditor.selection.paragraphFormat.rightIndent= 24;
    }

    public onPreview(): void {

        // this.container.documentEditor.saveAsBlob('Docx').then((exportedDocument: Blob) => {
        //     // The blob can be processed further
        //     let formData: FormData = new FormData();
        //     formData.append('fileName', 'sample.docx');
        //     formData.append('data', exportedDocument);
        //     console.log(formData);
        // });

        // this.container.documentEditor.saveAsBlob('Sfdt').then((exportedDocument: Blob) => {
        //     // The blob can be processed further
        //     // let formData: FormData = new FormData();
        //     // formData.append('fileName', 'sample.sfdt');
        //     // formData.append('data', exportedDocument);
        //     // console.log(formData);
        //     console.log(exportedDocument);

        //     // let sfdt: string = exportedDocument.stream;
        //     // this.documentEditor.open(sfdt);
        //   });

        let dataJson: string = this.container.documentEditor.serialize(); // note: data json
        
        let contentSurat = [];

        var kopSuratJson = defaultKopSurat;                     // note: kop surat
        let outputKopSuratJson = kopSuratJson["sections"];

        for (let i = 0; i < outputKopSuratJson.length; i++) {
            let outputBlocksKopSurat = outputKopSuratJson[i]["blocks"];
            for (let j = 0; j < outputBlocksKopSurat.length; j++) {
                contentSurat.push(outputBlocksKopSurat[j]);
            }
        }

        //---------------------------------------------------------------------------------------

        var contentJson = JSON.parse(dataJson);                 // note: content
        let outputContentJson = contentJson["sections"];

        for (let i = 0; i < outputContentJson.length; i++) {
            var outputBlocksContent = outputContentJson[i]["blocks"];
            for (let j = 0; j < outputBlocksContent.length; j++) {
                contentSurat.push(outputBlocksContent[j]);
            }
        }

        //---------------------------------------------------------------------------------------

        var ttdJson = defaultTtdSurat;                 // note: ttd
        let outputTtdJson = ttdJson["sections"];

        for (let i = 0; i < outputTtdJson.length; i++) {
            var outputBlocksTtd = outputTtdJson[i]["blocks"];
            for (let j = 0; j < outputBlocksTtd.length; j++) {
                contentSurat.push(outputBlocksTtd[j]);
            }
        }

        //--------------------------------------------------------------------------------------

        var resultJson = JSON.parse(dataJson);;                    //note: output
        let outputResultJson = resultJson["sections"];

        for (let i = 0; i < outputResultJson.length; i++) {
            outputResultJson[i]["blocks"] = contentSurat;
        }

        //----------------------------------------------------------------------------------------

        this.documentViewer.documentEditor.open(JSON.stringify(resultJson));
        //this.documentViewer.documentEditor.open(dataJson);

        this.container.documentEditor.focusIn();

        this.documentViewer.documentEditor.saveAsBlob('Docx').then((exportedDocument: Blob) => {
            // note:  convert to html
            convertToHtml({ arrayBuffer: exportedDocument })
                .then(function (result) {
                    var html = result.value; // The generated HTML
                    var messages = result.messages; // Any messages, such as warnings during conversion
                    console.log(html);
                    console.log(messages);
                }).done();
        });
       
    }

    onPreview2() {

        let dataJson: string = this.container.documentEditor.serialize(); // note: data json
        
        let contentSurat = [];

        var kopSuratJson = defaultKopSuratTanpaLogo;                     // note: kop surat
        let outputKopSuratJson = kopSuratJson["sections"];

        for (let i = 0; i < outputKopSuratJson.length; i++) {
            let outputBlocksKopSurat = outputKopSuratJson[i]["blocks"];
            for (let j = 0; j < outputBlocksKopSurat.length; j++) {
                contentSurat.push(outputBlocksKopSurat[j]);
            }
        }

        //---------------------------------------------------------------------------------------

        var contentJson = JSON.parse(dataJson);                 // note: content
        let outputContentJson = contentJson["sections"];

        for (let i = 0; i < outputContentJson.length; i++) {
            var outputBlocksContent = outputContentJson[i]["blocks"];
            for (let j = 0; j < outputBlocksContent.length; j++) {
                contentSurat.push(outputBlocksContent[j]);
            }
        }

        //---------------------------------------------------------------------------------------

        var ttdJson = defaultTtdSurat;                 // note: ttd
        let outputTtdJson = ttdJson["sections"];

        for (let i = 0; i < outputTtdJson.length; i++) {
            var outputBlocksTtd = outputTtdJson[i]["blocks"];
            for (let j = 0; j < outputBlocksTtd.length; j++) {
                contentSurat.push(outputBlocksTtd[j]);
            }
        }

        //--------------------------------------------------------------------------------------

        var resultJson = JSON.parse(dataJson);;                    //note: output
        let outputResultJson = resultJson["sections"];

        for (let i = 0; i < outputResultJson.length; i++) {
            outputResultJson[i]["blocks"] = contentSurat;
        }

        //----------------------------------------------------------------------------------------

        this.documentViewer.documentEditor.open(JSON.stringify(resultJson));
        //this.documentViewer.documentEditor.open(dataJson);

        this.container.documentEditor.focusIn();

        this.documentViewer.documentEditor.saveAsBlob('Docx').then((exportedDocument: Blob) => {
            // note:  convert to html
            convertToHtml({ arrayBuffer: exportedDocument })
                .then(function (result) {
                    var html = result.value; // The generated HTML
                    var messages = result.messages; // Any messages, such as warnings during conversion
                    console.log(html);
                    console.log(messages);
                }).done();
        });

    }

}