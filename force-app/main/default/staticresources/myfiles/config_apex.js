var resourceURL = '/resource/'
window.CoreControls.forceBackendType('ems');

var urlSearch = new URLSearchParams(location.hash)

//get custom object created in pdftronWvInstance.js and parse from iframe URL parameters
var custom = JSON.parse(urlSearch.get('custom'));
resourceURL = resourceURL + custom.namespacePrefix + 'V810';

/**
 * The following `window.CoreControls.set*` functions point WebViewer to the
 * optimized source code specific for the Salesforce platform, to ensure the
 * uploaded files stay under the 5mb limit
 */

// content edit workers
window.CoreControls.ContentEdit.setWorkerPath(resourceURL + 'content_edit');
window.CoreControls.ContentEdit.setResourcePath(resourceURL + 'content_edit_resource');

// office workers
window.CoreControls.setOfficeWorkerPath(resourceURL + 'office')
window.CoreControls.setOfficeAsmPath(resourceURL + 'office_asm');
window.CoreControls.setOfficeResourcePath(resourceURL + 'office_resource');

// pdf workers
window.CoreControls.setPDFResourcePath(resourceURL + 'resource')
if (custom.fullAPI) {
  window.CoreControls.setPDFWorkerPath(resourceURL + 'pdf_full')
  window.CoreControls.setPDFAsmPath(resourceURL + 'asm_full');
} else {
  window.CoreControls.setPDFWorkerPath(resourceURL + 'pdf_lean')
  window.CoreControls.setPDFAsmPath(resourceURL + 'asm_lean');
}

// external 3rd party libraries
window.CoreControls.setExternalPath(resourceURL + 'external')

//enable content editing feature
instance.UI.enableFeatures(instance.Feature.ContentEdit);

let currentDocId;

async function saveDocument() {
  // SF document file size limit
  const docLimit = 5 * Math.pow(1024, 2);
  const doc = instance.Core.documentViewer.getDocument();
  if (!doc) {
    return;
  }
  instance.openElement('loadingModal');
  const fileSize = await doc.getFileSize();
  const fileType = doc.getType();
  const filename = doc.getFilename();
  const xfdfString = await instance.Core.documentViewer.getAnnotationManager().exportAnnotations();
  const data = await doc.getFileData({
    // Saves the document with annotations in it
    xfdfString
  });

  let binary = '';
  const bytes = new Uint8Array(data);
  for (let i = 0; i < bytes.byteLength; i++) {
    binary += String.fromCharCode(bytes[i]);
  }

  const base64Data = window.btoa(binary);

  const payload = {
    title: filename.replace(/\.[^/.]+$/, ""),
    filename,
    base64Data,
    contentDocumentId: currentDocId
  }
  // download files larger than 5MB, else save back to Salesforce files
  fileSize < docLimit ? parent.postMessage({ type: 'SAVE_DOCUMENT', payload }, '*') : downloadWebViewerFile();
}

//create a custom modal
const createSavedModal = (instance) => {
  const divInput = document.createElement('div');
  divInput.innerText = 'File saved successfully.';
  const modal = {
    dataElement: 'savedModal',
    body: {
      className: 'myCustomModal-body',
      style: {
        'text-align': 'center'
      },
      children: [divInput]
    }
  }
  instance.UI.addCustomModal(modal);
}

//get the current file data from WebViewer and download as blob
const downloadWebViewerFile = async () => {
  const doc = instance.Core.documentViewer.getDocument();

  if (!doc) {
    return;
  }

  const data = await doc.getFileData();
  const arr = new Uint8Array(data);
  const blob = new Blob([arr], { type: 'application/pdf' });

  const filename = doc.getFilename();

  downloadFile(blob, filename)
}

//download a blob manually
const downloadFile = (blob, fileName) => {
  const link = document.createElement('a');
  // create a blobURI pointing to our Blob
  link.href = URL.createObjectURL(blob);
  link.download = fileName;
  // some browser needs the anchor to be in the doc
  document.body.append(link);
  link.click();
  link.remove();
  // in case the Blob uses a lot of memory
  setTimeout(() => URL.revokeObjectURL(link.href), 7000);
};

//on viewer load, execute the following logic
window.addEventListener('viewerLoaded', async function () {
  //select Edit ribbon
  instance.setToolbarGroup(instance.UI.ToolbarGroup.EDIT);

  //auto select a tool from https://www.pdftron.com/api/web/Core.Tools.html#.ToolNames__anchor
  instance.UI.setToolMode(instance.Core.Tools.ToolNames.TEXT_SELECT);

  //add hotkeys to execute saveDocument() function
  instance.hotkeys.on('ctrl+s, command+s', e => {
    e.preventDefault();
    saveDocument();
  });

  // Create a button, with a disk icon, to invoke the saveDocument function
  instance.setHeaderItems(function (header) {
    var myCustomButton = {
      type: 'actionButton',
      dataElement: 'saveDocumentButton',
      title: 'tool.SaveDocument',
      img: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path d="M0 0h24v24H0z" fill="none"/><path d="M17 3H5c-1.11 0-2 .9-2 2v14c0 1.1.89 2 2 2h14c1.1 0 2-.9 2-2V7l-4-4zm-5 16c-1.66 0-3-1.34-3-3s1.34-3 3-3 3 1.34 3 3-1.34 3-3 3zm3-10H5V5h10v4z"/></svg>',
      onClick: function () {
        saveDocument();
      }
    }
    header.get('viewControlsButton').insertBefore(myCustomButton);
  });

  // When the viewer has loaded, this makes the necessary call to get the
  // pdftronWvInstance code to pass User Record information to this config file
  // to invoke annotManager.setCurrentUser
  instance.Core.documentViewer.getAnnotationManager().setCurrentUser(custom.username);

  //add the custom modal to WebViewer UI
  createSavedModal(instance);
});

//on completion of document load, execute logic
window.addEventListener('documentLoaded', () => {
  //your function here
})

//on receive of payload from postMessage call, execute receiveMessage() function
window.addEventListener("message", receiveMessage, false);

//handler for postMessage()
function receiveMessage(event) {
  if (event.isTrusted && typeof event.data === 'object') {
    switch (event.data.type) {
      case 'OPEN_DOCUMENT':
        instance.loadDocument(event.data.file, {
          officeOptions: {
            disableBrowserFontSubstitution: true,
          }
        })
        break;
      case 'OPEN_DOCUMENT_BLOB':
        const { blob, extension, filename, documentId } = event.data.payload;
        console.log("documentId", documentId);
        currentDocId = documentId;
        instance.loadDocument(blob, { extension, filename, documentId })
        break;
      case 'DOCUMENT_SAVED':
        console.log(`${JSON.stringify(event.data)}`);
        instance.UI.openElements(['savedModal']);
        setTimeout(() => {
          instance.closeElements(['savedModal', 'loadingModal'])
        }, 2000)
        break;
      case 'LMS_RECEIVED':  
        instance.loadDocument(event.data.payload.message, {
          filename: event.data.payload.filename,
          withCredentials: false
        });
        break;
      case 'DOWNLOAD_DOCUMENT':
        downloadWebViewerFile();
        break;
      case 'CLOSE_DOCUMENT':
        instance.closeDocument()
        break;
      default:
        break;
    }
  }
}