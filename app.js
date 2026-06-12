const sampleCustomXml = `<Properties xmlns="http://schemas.officeatwork.com/2022/templateProperties">
<officeatwork_languages>
  <Value>en</Value>
  <Value>de</Value>
</officeatwork_languages>
<subject>Invitation to branch opening</subject>
<subject.de>Einladung zur Geschäftsstelleneröffnung</subject.de>
<location>
  <city>Zug</city>
  <country>Switzerland</country>
  <country.de>Schweiz</country.de>
  <street>Bundesplatz 12</street>
</location>
</Properties>`;

const sampleDocumentProperties = [
    {
        type: 'customdocumentproperty',
        name: 'OutlookScenario',
        value: 'Template Chooser embedded dialog spike',
    },
    {
        type: 'customdocumentproperty',
        name: 'CustomerReference',
        value: 'TC-Embedded-Outlook',
    },
    {
        type: 'builtindocumentproperty',
        name: 'company',
        value: 'officeatwork',
    },
];

const elements = {
    customXml: document.querySelector('#customXml'),
    documentLink: document.querySelector('#documentLink'),
    documentProperties: document.querySelector('#documentProperties'),
    openButton: document.querySelector('#openTemplateChooser'),
    result: document.querySelector('#result'),
    templateChooserUrl: document.querySelector('#templateChooserUrl'),
};

elements.customXml.value = sampleCustomXml;
elements.documentProperties.value = JSON.stringify(
    sampleDocumentProperties,
    null,
    2,
);

Office.onReady(officeInfo => {
    const displayDialogAsync = Office.context?.ui?.displayDialogAsync;
    const supportsDialogApi =
        Office.context?.requirements?.isSetSupported('Mailbox', '1.4') ?? false;

    if (!officeInfo.host) {
        showError(
            new Error(
                'Office.js loaded, but this page is not running inside an Office add-in. Open it from the sideloaded Outlook ribbon command.',
            ),
        );
        return;
    }

    if (!supportsDialogApi || typeof displayDialogAsync !== 'function') {
        showError(
            new Error(
                `Dialog API is unavailable. Host: ${officeInfo.host}; platform: ${officeInfo.platform}; Mailbox 1.4: ${supportsDialogApi}.`,
            ),
        );
        return;
    }

    elements.result.textContent = `Office.js ready: ${officeInfo.host} (${officeInfo.platform})`;
    elements.openButton.disabled = false;
    elements.openButton.addEventListener('click', openTemplateChooser);
});

function openTemplateChooser() {
    clearResult();

    let dialogUrl;
    try {
        dialogUrl = createDialogUrl();
    } catch (error) {
        showError(error);
        return;
    }

    const displayDialogAsync = Office.context?.ui?.displayDialogAsync;
    if (typeof displayDialogAsync !== 'function') {
        showError(new Error('Office Dialog API is no longer available.'));
        return;
    }

    displayDialogAsync.call(
        Office.context.ui,
        dialogUrl,
        { displayInIframe: false, height: 80, width: 80 },
        result => {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
                showError(result.error);
                return;
            }

            const dialog = result.value;
            dialog.addEventHandler(
                Office.EventType.DialogMessageReceived,
                message => {
                    dialog.close();
                    showMessage(message.message);
                },
            );
            dialog.addEventHandler(
                Office.EventType.DialogEventReceived,
                event => {
                    if (event.error !== 12006) {
                        showError(event);
                    }
                },
            );
        },
    );
}

function createDialogUrl() {
    const documentProperties = JSON.parse(elements.documentProperties.value);
    if (!Array.isArray(documentProperties)) {
        throw new Error('Document properties must be a JSON array.');
    }

    const injectedData = [
        {
            type: 'customxmlpart',
            base64Xml: encodeBase64(elements.customXml.value),
        },
        ...documentProperties,
    ];

    const targetUrl = new URL(elements.templateChooserUrl.value);
    targetUrl.searchParams.set(
        'inject',
        encodeBase64(JSON.stringify(injectedData)),
    );
    targetUrl.searchParams.set('dialogParentOrigin', location.origin);
    targetUrl.hash = '/template-chooser-embedded';

    const redirectUrl = new URL('dialog-redirect.html', location.href);
    redirectUrl.searchParams.set('target', targetUrl.toString());
    return redirectUrl.toString();
}

function showMessage(message) {
    let parsedMessage;
    try {
        parsedMessage = JSON.parse(message);
    } catch {
        elements.result.textContent = message;
        return;
    }

    elements.result.textContent = JSON.stringify(parsedMessage, null, 2);

    if (
        parsedMessage.type === 'template-chooser-document-created' &&
        parsedMessage.fileAsBase64
    ) {
        const bytes = Uint8Array.from(atob(parsedMessage.fileAsBase64), char =>
            char.charCodeAt(0),
        );
        const blob = new Blob([bytes], {
            type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        });
        const fileName = parsedMessage.fileName || 'document.docx';

        elements.documentLink.href = URL.createObjectURL(blob);
        elements.documentLink.download = fileName;
        elements.documentLink.textContent = fileName;
        elements.documentLink.hidden = false;
    }
}

function clearResult() {
    elements.documentLink.hidden = true;
    elements.documentLink.removeAttribute('href');
    elements.result.textContent = 'Waiting for Template Chooser';
}

function showError(error) {
    elements.result.textContent =
        error instanceof Error ? error.message : JSON.stringify(error, null, 2);
}

function encodeBase64(value) {
    const bytes = new TextEncoder().encode(value);
    let binary = '';
    bytes.forEach(byte => {
        binary += String.fromCharCode(byte);
    });
    return btoa(binary);
}
