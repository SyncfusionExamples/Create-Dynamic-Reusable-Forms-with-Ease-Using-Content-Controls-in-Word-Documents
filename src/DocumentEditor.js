import { useEffect, useRef, useState, useCallback,useMemo} from "react";
import {
  DocumentEditorContainerComponent,
  Toolbar,
} from "@syncfusion/ej2-react-documenteditor";
import TitleBar from "./titlebar";
import contentControlData from './ContentControlElements.json';

DocumentEditorContainerComponent.Inject(Toolbar);

/**
 * Tracks elements that have already had drag event listeners attached.
 * Prevents duplicate listeners and unnecessary state updates during re-renders.
 */
export let attachedDragElements = new Set();

function DocumentEditor() {
  const containerRef = useRef(null);
  const titleBarRef = useRef(null);
  const [isPropertyDialog, setPropertyDialog] = useState(false);
  const [contentControlElementCount, setContentControlElementCount] = useState(0);
  let isDragElement = false;
  let currentContentControl = null;
  const defaultDocument = '';
  let contentControlList = JSON.parse(localStorage.getItem("contentControlList")) || contentControlData;
  
  
  useEffect(() => {
    initializeEditor();
    setupContentControlInsertion();
    setupDragAndDrop();
  }, []);

  useEffect(() => {
    setupDragAndDrop();
  }, [isPropertyDialog, contentControlElementCount]);

  /**
   * Initializes the document editor and sets up the title bar.
   */
  function initializeEditor() {
    if (containerRef.current && !titleBarRef.current) {
      convertDocxToSfdt();
      titleBarRef.current = new TitleBar(
        document.getElementById("documenteditor_titlebar"),
        containerRef.current.documentEditor,
        true
      );
      //Update the content control list in local storage
      let contentControlElement = JSON.parse(localStorage.getItem("contentControlList"));
      if (!contentControlElement) {
        localStorage.setItem("contentControlList", JSON.stringify(contentControlList));
      }
      containerRef.current.documentEditor.documentName = "Customer Review Form";
      titleBarRef.current.updateDocumentTitle();
      containerRef.current.documentChange = () => {
        titleBarRef.current?.updateDocumentTitle();
        containerRef.current?.documentEditor.focusIn();
      };
    }
  }

    // Convert GitHub Raw document to SFDT and load in Editor.
  const convertDocxToSfdt = async () => {
    try {
      const docxResponse = await fetch('https://raw.githubusercontent.com/syncfusion/blazor-showcase-document-explorer/master/server/wwwroot/Files/Documents/Giant%20Panda.docx');
      const docxBlob = await docxResponse.blob();

      const formData = new FormData();
      formData.append('files', docxBlob, 'GiantPanda.docx');

      const importResponse = await fetch('https://ej2services.syncfusion.com/production/web-services/api/documenteditor/Import', {
        method: 'POST',
        body: formData,
      });

      if (importResponse.ok) {
        defaultDocument = await importResponse.text();
        containerRef.current.documentEditor.open(defaultDocument);
      } else {
        console.error(`Failed to import document: ${importResponse.statusText}`);
      }
    } catch (error) {
      console.error('Error converting document:', error);
    }
  };

  /**
   * Sets up drag-and-drop functionality for content control elements.
   */
  function setupDragAndDrop() {
    const container = document.getElementById("container");
    const deleteElement = document.querySelectorAll('.e-trash');
    document.querySelectorAll(".content-control").forEach(element => {
      if (!attachedDragElements.has(element)) {
        element.addEventListener("dragstart", (event) => {
          event.dataTransfer.setData("controlType", event.target.dataset.type);
          deleteElement.forEach(icon => icon.style.display = 'none');
          isDragElement = true;
        });
        attachedDragElements.add(element);
      }
    });
    container?.addEventListener("dragover", (event) => event.preventDefault());
    container?.addEventListener("drop", handleDrop);
  }

  /**
   * Handles the drop event when a content control is dragged into the editor.
   */
  function handleDrop(event) {
    event.preventDefault();
    const deleteElement = document.querySelectorAll('.e-trash');
    const type = event.dataTransfer.getData("controlType");
    deleteElement.forEach(icon => icon.style.display = 'block');
    containerRef.current.documentEditor.selection.select({
      x: event.offsetX,
      y: event.offsetY,
      extend: false,
    });
    if (!type || !containerRef.current || !isDragElement) return;
    const editor = containerRef.current.documentEditor.editor;
    const control = contentControlList[type];
    if (control) {
      editor.insertContentControl(control);
      let contentControlDetails = editor.selection.getContentControlInfo();
      if (contentControlDetails) {
        if (currentContentControl && (contentControlDetails.type === "Text" || contentControlDetails.type === "RichText")) {
          currentContentControl[0].contentControlProperties.hasPlaceHolderText = true;
        } else if (currentContentControl && contentControlDetails.type === "CheckBox" && !control.canEdit) {
          currentContentControl[0].contentControlProperties.lockContents = false;
        }
      }
    }
    isDragElement = false;
  }

  /**
   * Generates a list of content controls from the control content list for rendering in the sidebar.
   */
  const controlList = Object.entries(contentControlList).map(([type, config]) => ({
    label: config.title,
    tag: config.tag,
    type,
  }));

   /**
   * Sets up event for inserting content controls into the document editor.
   * This function ensures that only valid content controls with proper title and tag
   * are added to the sidebar, and updates the control map accordingly.
   */
  const setupContentControlInsertion = useCallback(() => {
    const editor = containerRef.current?.documentEditor;
    if (!editor) return;
    containerRef.current.toolbarClick = function (args) {
      let contentControls = ["RichTextContentControl", "PlainTextContentControl", "ComboBoxContentControl", "DropDownListContentControl", "CheckBoxContentControl", "DatePickerContentControl"];
      let insertContentControl = args.item.id === "container_toolbar_content_control";
      let isDialogOpen = false;
      editor.contentChange = function (args) {
        let currentContentControlType = editor.selection.contextType;
        if (contentControls.includes(currentContentControlType)) {
          let contentControlDetails = editor.selection.getContentControlInfo();
          if (contentControlDetails && insertContentControl && (contentControlDetails.title === undefined || contentControlDetails.tag === undefined)) {
            editor.showDialog('ContentControlProperties');
            setPropertyDialog(true);
            isDialogOpen = true;
            insertContentControl = false;
          }
          if (contentControlDetails && isDialogOpen && (contentControlDetails.title !== undefined && contentControlDetails.title !== "" && contentControlDetails.title !== undefined && contentControlDetails.tag !== "")) {
            isDialogOpen = false;
            setPropertyDialog(false);
            let listItem = [];
            if (currentContentControlType === "ComboBoxContentControl" || currentContentControlType === "DropDownListContentControl") {
              const listValues = args.source.contentControlPropertiesDialogModule.listviewInstance.curViewDS;
              const uniqueValues = [
                ...new Set(
                  listValues
                    .map(item => item.value)
                    .filter(value => value !== 'Choose an item')
                )
              ];
              listItem = [...uniqueValues];
            }
            const formattedTag = contentControlDetails.tag.trim().toLowerCase().replace(/\s+/g, '_');
            contentControlList[formattedTag] = {
              title: contentControlDetails.title,
              tag: contentControlDetails.tag,
              value: currentContentControlType === "CheckBoxContentControl" ? false : contentControlDetails.tag,
              type: contentControlDetails.type,
              items: listItem,
              canDelete: currentContentControlType === "ComboBoxContentControl" || currentContentControlType === "CheckBoxContentControl" ? contentControlDetails.canDelete : !contentControlDetails.canDelete,
              canEdit: currentContentControlType === "ComboBoxContentControl" || currentContentControlType === "CheckBoxContentControl" ? contentControlDetails.canEdit : !contentControlDetails.canEdit,
            };
            if (currentContentControlType !== "CheckBoxContentControl") {
              containerRef.current.documentEditor.editor.insertText(contentControlDetails.tag);
            }
            // Update the local storage with the updated list
            localStorage.setItem("contentControlList", JSON.stringify(contentControlList));
          }
        }
      }
    }
    editor.selectionChange = (args) => {
      if(args.source.selection.contentControls){
        currentContentControl = args.source.selection.contentControls
      }
    };
  }, []);

  /**
   * Removes a content control from the control control list by its tag.
   */
  function removeContentControl(tag) {
    const tagValue = tag.trim().toLowerCase().replace(/\s+/g, '_');
    const keyToRemove = Object.keys(contentControlList).find(key => {
      const item = contentControlList[key];
      const formattedValue = item.tag.trim().toLowerCase().replace(/\s+/g, '_');
      return formattedValue === tagValue;
    });
    if (keyToRemove) {
      delete contentControlList[keyToRemove];
      setContentControlElementCount(prev => prev + 1);
      // Update the local storage with the updated list
      localStorage.setItem("contentControlList", JSON.stringify(contentControlList));
    }
  }

  return (
    <div id="mainContainer">
      <div id="documenteditor_titlebar" className="e-de-ctn-title"></div>
      <div className="control-pane">
        <div className="content-control-panel">
          <h4>Select Field to Insert</h4>
          {controlList.map((item, index) => (
            <div className="content-controls" key={`${item.tag}_${index}`}>
              <div
                className="content-control"
                data-type={item.type}
                draggable
              >
                <span>{item.label}</span>
                <button
                  className="e-icons e-trash"
                  onClick={() => removeContentControl(item.tag)}
                >
                </button>
              </div>
            </div>
          ))}
        </div>
        <div className="control-section">
          <DocumentEditorContainerComponent
            id="container"
            ref={containerRef}
            height={"calc(100vh - 40px)"}
            serviceUrl="https://services.syncfusion.com/react/production/api/documenteditor/"
            enableToolbar={true}
          />
        </div>
      </div>
    </div>
  );
}

export default DocumentEditor;
