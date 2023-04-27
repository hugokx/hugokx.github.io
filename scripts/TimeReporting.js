﻿(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            loadProjects();
            loadPrestations();
            document.getElementById('inputForm').addEventListener('submit', handleSubmit)
        });
    };

    function loadProjects() {
        fetch('/files/projets.csv')
            .then(response => {
                if (response.ok) {
                    return response.text();
                } else {
                    throw new Error('Failed to load the CSV file');
                }
            })
            .then(data => {
                const lines = data.split('\n');
                const projectSelect = document.getElementById('project');

                lines.forEach(line => {
                    const [name, description] = line.split(';');
                    const option = document.createElement('option');
                    option.value = name;
                    option.textContent = name;
                    option.title = description;
                    projectSelect.appendChild(option);
                });
            })
            .catch(error => {
                console.error('Error:', error);
            });
    }
    function loadPrestations() {
        fetch('/files/prestations.csv')
            .then(response => {
                if (response.ok) {
                    return response.text();
                } else {
                    throw new Error('Failed to load the CSV file');
                }
            })
            .then(data => {
                const lines = data.split('\n');
                const prestationSelect = document.getElementById('prestation');

                lines.forEach(line => {
                    const name = line.trim();
                    const option = document.createElement('option');
                    option.value = name;
                    option.textContent = name;

                    if (name === 'N/A') {
                        option.selected = true;
                    }
                    prestationSelect.appendChild(option);
                });
            })
            .catch(error => {
                console.error('Error:', error);
            });
    }

    function getOutlookClientType() {
        const hostName = Office.context.mailbox.diagnostics.hostName;
      
        if (hostName === 'Outlook') {
          return 'Outlook Desktop';
        } else if (hostName === 'OutlookWebApp') {
          return 'Outlook Web App';
        } else {
          return 'Unknown';
        }
      }
    
    function removeSpellAndGramTags(text) {
        return text.replace(/<span class=(['"])?(?:SpellE|GramE)\1?>(.*?)<\/span>/g, '');
    }

    function removeSpellCheckTags(htmlString) {
        const regex = /<span class=(['"])?(?:SpellE|GramE)\1?>(.*?)<\/span>/g;
        return htmlString.replace(regex, (match, quote, content) => {
            console.log('Matched:', match);
            console.log('Content:', content);
            return content;
          });
    }

    function handleSubmit(event) {
        event.preventDefault(); // Prevent the default form submission behavior

        const projectSelect = document.getElementById('project');
        const paeProjectInput = document.getElementById('pae_project');
        const prestationSelect = document.getElementById('prestation');
        const includeCheck = document.getElementById('notInclude');

        const projectType = projectSelect.options[projectSelect.selectedIndex].textContent;
        const paeProjectType = paeProjectInput.value;
        const prestationType = prestationSelect.options[prestationSelect.selectedIndex].textContent;
        const includeValue = includeCheck.value;

        const customReport = `{projet:${projectType};projet_pae:${paeProjectType};prestation:${prestationType};inclu:${includeValue}}`;
        const customText = `-----------------------------------------------------<br><span style="color: white;">${customReport}</span>`;

        const clientType = getOutlookClientType();

        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                let currentDescription = result.value;
                console.log("---------START OF CURRENT DESCRIPTION---------");
                console.log(currentDescription);
                console.log("---------END OF CURRENT DESCRIPTION---------");
                const reportRegex = /-----------------------------------------------------\s*<br>\s*<span style="color: white;">([\s\S]*?)<\/span>/;

                currentDescription = removeSpellCheckTags(currentDescription);
                console.log("---------START OF CURRENT DESCRIPTION WITHOUT GRAMMAR TAG---------");
                console.log(currentDescription);
                console.log("---------END OF CURRENT DESCRIPTION WITHOUT GRAMMAR TAG---------");

                if (reportRegex.test(currentDescription)) {
                    const currentReport = reportRegex.exec(currentDescription)[1];

                    if (currentReport === customReport) {
                        showAlertDialog("Les éléments du reporting sont déjà présents");
                    } else {
                        const [_, projet_tmp, pae_tmp, prestation_tmp, include_tmp] = currentReport.match(/\{projet:(.*?);projet_pae:(.*?);prestation:(.*?);inclu:(.*?)\}/);
                        const dialogContent = `
                                            Un reporting existe déjà avec les éléments suivants :<br>
                                            Projet: ${projet_tmp}<br>
                                            Projet PAE: ${pae_tmp}<br>
                                            Type de prestation : ${prestation_tmp}<br><br>
                                            Voulez-vous le remplacer?
                                            `;

                        showConfirmationDialog(dialogContent, () => {
                            // User confirmed, replace the line
                            const updatedDescription = currentDescription.replace(currentReport, customReport);
                            updateEventDescription(updatedDescription);
                        }, () => {
                            // User canceled, don't change the description
                            console.log('User canceled, description not changed.');
                        });
                    }
                } else {
                    let updatedDescription;
                    
                    if (clientType === 'Outlook Desktop') {
                        const insertionPoint = '</div></body></html>';
                        const index = currentDescription.lastIndexOf(insertionPoint);
                        updatedDescription = currentDescription.slice(0, index) + '<div>' + customText + '</div>' + currentDescription.slice(index);
                    } else if (clientType === 'Outlook Web App') {
                        updatedDescription = currentDescription + '<div>' + customText + '</div>';
                    } else {
                        console.error('Unknown client type'); 
                        return;
                    }
                    updateEventDescription(updatedDescription);
                    showAlertDialog("Les éléments ont bien été ajoutés");
                }
            } else {
                console.error('Failed to get the current event description:', result.error);
            }
        });
    }

    function updateEventDescription(updatedDescription) {
        Office.context.mailbox.item.body.setAsync(updatedDescription, { coercionType: Office.CoercionType.Html }, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log('Event description updated successfully.');
            } else {
                console.error('Failed to update the event description:', result.error);
            }
        });
    }

    function showConfirmationDialog(message, onConfirm, onCancel) {
        const confirmationDialogContent = document.getElementById('confirmationDialogContent');
        confirmationDialogContent.innerHTML = message;

        const confirmationDialog = new bootstrap.Modal(document.getElementById('confirmationDialog'), {});
        const confirmButton = document.getElementById('confirmButton');
        const cancelButton = document.getElementById('cancelButton');

        confirmButton.addEventListener('click', function () {
            onConfirm();
            confirmationDialog.hide();
        });

        cancelButton.addEventListener('click', function () {
            onCancel();
            confirmationDialog.hide();
        });

        confirmationDialog.show();
    }

    function showAlertDialog(message) {
        const alertModalBody = document.getElementById('alertModalBody');
        alertModalBody.innerHTML = message;

        var alertModal = new bootstrap.Modal(document.getElementById('alertModal'), {});
        alertModal.show();
    }

    /**function showConfirmationDialog(message, onConfirm, onCancel) {
        const dialogHTML = `
        <div class="ms-Dialog">
            <div class="ms-Dialog-title">Attention</div>
            <div class="ms-Dialog-content">
                <p class="ms-Dialog-subText">${message}</p>
            </div>
            <div class="ms-Dialog-actions">
                <button class="ms-Dialog-action ms-Button ms-Button--primary"><span class="ms-Button-label">Oui</span></button>
                <button class="ms-Dialog-action ms-Button"><span class="ms-Button-label">Non</span></button>
            </div>
        </div>
    `;

        const container = document.createElement('div');
        container.innerHTML = dialogHTML;
        document.body.appendChild(container);

        const dialog = new fabric['Dialog'](container.querySelector('.ms-Dialog'));
        dialog.open();

        const primaryButton = container.querySelector('.ms-Button--primary');
        primaryButton.addEventListener('click', () => {
            dialog.close();
            document.body.removeChild(container);
            onConfirm();
        });

        const secondaryButton = container.querySelector('.ms-Button:not(.ms-Button--primary)');
        secondaryButton.addEventListener('click', () => {
            dialog.close();
            document.body.removeChild(container);
            onCancel();
        });
    }**/


})();