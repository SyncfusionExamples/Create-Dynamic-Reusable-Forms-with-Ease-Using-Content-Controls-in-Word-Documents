import { createElement } from '@syncfusion/ej2-base';
import { Button } from '@syncfusion/ej2-buttons';
import { DropDownButton } from '@syncfusion/ej2-splitbuttons';

/**
 * Represents the title bar of the document editor.
 */
class TitleBar {
    constructor(element, docEditor, isShareNeeded, dialogComponent) {
        // Initialize title bar elements.
        this.tileBarDiv = element;
        this.documentEditor = docEditor;
        this.dialogComponent = dialogComponent;
        this.initializeTitleBar(isShareNeeded);
        this.wireEvents();

        // Append document title after initialization
        if (this.tileBarDiv && !this.tileBarDiv.contains(this.documentTitle)) {
            this.tileBarDiv.prepend(this.documentTitle);
        }
    }

    /**
     * Initializes the title bar, setting up text and tooltips.
     * @param {boolean} isShareNeeded - Flag to decide if sharing feature is needed.
     */
    initializeTitleBar = (isShareNeeded) => {
        const downloadText = 'Download';
        const downloadToolTip = 'Download this document.';
        const printText = 'Print';
        const printToolTip = 'Print this document (Ctrl+P).';

        // Create the document title element
        this.documentTitle = createElement('label', {
            id: 'documenteditor_title_name',
            styles: 'font-weight:400;text-overflow:ellipsis;white-space:pre;overflow:hidden;user-select:none;cursor:text; font-size:16px'
        });
        this.documentTitle.innerHTML = this.documentEditor.documentName;

        const btnStyles = 'float:right;background: transparent;box-shadow:none; font-family: inherit;border-color: transparent;' +
            'border-radius: 2px;color:inherit;font-size:12px;text-transform:capitalize;height:28px;font-weight:400;margin: 4px;';

        // Initialize print button
        this.print = this.addButton('e-icons e-print e-de-padding-right', printText, btnStyles, 'de-print', printToolTip, false);

        // Initialize download drop-down button
        const items = [
            { text: 'Syncfusion® Document Text (*.sfdt)', id: 'sfdt' },
            { text: 'Word Document (*.docx)', id: 'word' },
            { text: 'Word Template (*.dotx)', id: 'dotx' },
            { text: 'Plain Text (*.txt)', id: 'txt' },
        ];
        this.download = this.addButton('e-icons e-download e-de-padding-right', downloadText, btnStyles, 'documenteditor-share', downloadToolTip, true, items);

        // Hide download button if sharing is not needed
        if (!isShareNeeded) {
            this.download.element.style.display = 'none';
        }
    };

    /**
     * Wires events to the buttons.
     */
    wireEvents = () => {
        this.print.element.addEventListener('click', this.onPrint);
    };

    /**
     * Updates the document title displayed in the title bar.
     */
    updateDocumentTitle = () => {
        if (this.documentEditor.documentName === '') {
            this.documentEditor.documentName = 'Untitled';
        }
        this.documentTitle.textContent = this.documentEditor.documentName;
    };

    /**
     * Adds a button to the title bar.
     * @param {string} iconClass - Icon class for the button.
     * @param {string} btnText - Button text.
     * @param {string} styles - Button styles.
     * @param {string} id - Button id.
     * @param {string} tooltipText - Tooltip text for the button.
     * @param {boolean} isDropDown - Whether the button is a dropdown.
     * @param {Array} items - Items for dropdown, if applicable.
     * @returns {Button|DropDownButton} - The created button instance.
     */
    addButton(iconClass, btnText, styles, id, tooltipText, isDropDown, items) {
        // Create button element
        const button = createElement('button', { id, styles });
        this.tileBarDiv.appendChild(button);
        button.setAttribute('title', tooltipText);

        // Return appropriate button based on isDropDown flag
        if (isDropDown) {
            return new DropDownButton({
                select: this.onDownloadClick,
                items,
                iconCss: iconClass,
                cssClass: 'e-caret-hide',
                content: btnText,
            }, button);
        } else {
            return new Button({ iconCss: iconClass, content: btnText }, button);
        }
    }

    /**
     * Handles print button click event.
     */
    onPrint = () => {
        this.documentEditor.print();
    };

    /**
     * Handles download item selection.
     * @param {Object} args - Event arguments.
     */
    onDownloadClick = (args) => {
        const formatMap = {
            'word': 'Docx',
            'sfdt': 'Sfdt',
            'txt': 'Txt',
            'dotx': 'Dotx'
        };
        this.save(formatMap[args.item.id]);
    };

    /**
     * Saves the document in the specified format.
     * @param {string} format - Format to save the document as.
     */
    save = (format) => {
        this.documentEditor.save(this.documentEditor.documentName || 'sample', format);
    };
}

export default TitleBar;