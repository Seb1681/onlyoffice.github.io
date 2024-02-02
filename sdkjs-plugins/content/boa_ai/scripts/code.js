// an Chat plugin of AI
(function (window, undefined) {
    let token = '';
    let rsdId = '';

    window.Asc.plugin.init = async function ()
    {
        const uuid = uuidv4();
		const payload = {
			onlyOfficePlugin: 'GetAiMetaData',
			pluginId: uuid
		}
		const connection = new signalR.HubConnectionBuilder()
			.withUrl("https://boa-admin-vnext.dev.azaas.online/signalr-hubs/onlyOffice?pluginId=" + uuid, {
				skipNegotiation: true,
				transport: signalR.HttpTransportType.WebSockets
			  })
			.configureLogging(signalR.LogLevel.Information)
			.build();

		connection.on("ReceiveMessage", (user, message) => {
			console.log('Message on RSD Pilot: ' + message);
			if (message) {
                const initData = JSON.parse(message);
                token= initData.token;
                rsdId = initData.rsdId;
			    connection.stop().then(() => console.log('RSD Pilot connection successfully closed.')).catch(err => console.error('Error while closing the connection: ', err));
			}
		});
        console.log('RSD Pilot Plugin Initiated')
        await connection.start();
        console.log("SignalR Connected from RSD Pilot");
        window.parent.parent.postMessage(payload, '*');
    };

    window.Asc.plugin.button = function () {
        this.executeCommand("close", "");
    };

    function getContextMenuItems() {
        let settings = {
            guid: window.Asc.plugin.guid,
            items: [
                {
                    id: 'RSD Pilot',
                    text: generateText('RSD Pilot'),
                    items: [
                        {
                            id: 'generate',
                            text: generateText('Generate'),
                        },
                        {
                            id: 'rephrase',
                            text: generateText('Rephrase'),
                        },
                    ]
                }
            ]
        }
        return settings;
    }

    window.Asc.plugin.attachEvent('onContextMenuShow', function (options) {
        if (!options) return;

        if (options.type === 'Selection' || options.type === 'Target')
            this.executeMethod('AddContextMenuItem', [getContextMenuItems()]);
    });

    const parseMarkdown = (markdownString) => {
        const lines = markdownString.split('\n');
        const categorized = lines.map(line => {
        // Headings
        if (line.startsWith('# ')) {
            return { type: 'Heading 1', content: line.substring(2) };
        } else if (line.startsWith('## ')) {
            return { type: 'Heading 2', content: line.substring(3) };
        } else if (line.startsWith('### ')) {
            return { type: 'Heading 3', content: line.substring(4) };
        } else if (line.startsWith('#### ')) {
            return { type: 'Heading 4', content: line.substring(5) };
        } else if (line.startsWith('##### ')) {
            return { type: 'Heading 5', content: line.substring(6) };
        } 
        // Blockquotes
        // else if (line.startsWith('> ')) {
        //     return { type: 'blockquote', content: line.substring(2) };
        // }
        // // Ordered List
        // else if (/^\d+\.\s/.test(line)) {
        //     return { type: 'ordered_list_item', content: line.substring(line.indexOf(' ') + 1) };
        // }
        // // Unordered List
        // else if (/^[-*+]\s/.test(line)) {
        //     return { type: 'unordered_list_item', content: line.substring(2) };
        // }
        // // Code Blocks
        // else if (line.startsWith('```')) {
        //     return { type: 'code_block', content: line.substring(3) };
        // }
        // // Horizontal Rules
        // else if (/^---$|^___$|^\*\*\*$/.test(line)) {
        //     return { type: 'horizontal_rule', content: '' };
        // }
        // Default to paragraph if no other type matches
        else {
            return { type: 'Normal', content: line };
        }
    });

    return categorized;
    }

    // generate content in document
    window.Asc.plugin.attachContextMenuClickEvent('generate', function () {
        window.Asc.plugin.executeMethod('GetSelectedText', null, (text) => {
            window.parent.parent.postMessage({"onlyOfficePlugin": "Loading"}, '*');
            let prompt = (`Please generate the content for: "${text.trim()}". Please reply only the content, in markdown format.`);
            sseRequest(prompt)
                .then(result => {
                    Asc.scope.p = parseMarkdown(result.text);
                    Asc.plugin.callCommand(function () {
                        let oDocument = Api.GetDocument();
                        Asc.scope.p.forEach((item) => {
                            var oStyle = oDocument.GetStyle(item.type);
                            let oParagraph = Api.CreateParagraph();
                            oParagraph.SetStyle(oStyle);
                            oParagraph.AddText(item.content);
                            oDocument.InsertContent([oParagraph]);
                        });
                    })
                })
                .catch(error => {
                    console.error(error);
                })
                .finally(() => {
                    window.parent.parent.postMessage({"onlyOfficePlugin": "Loading"}, '*');
                });
            });
    });

    // rephrase content in document
    window.Asc.plugin.attachContextMenuClickEvent('rephrase', function () {
        window.Asc.plugin.executeMethod('GetSelectedText', null, (text) => {
            window.parent.parent.postMessage({"onlyOfficePlugin": "Loading"}, '*');
            let prompt = (`Please rephrase this sentence: "${text.trim()}". Please reply only the content, in markdown format.`);
            sseRequest(prompt)
                .then(result => {
                    Asc.scope.p = parseMarkdown(result.text);
                    Asc.plugin.callCommand(function () {
                        let oDocument = Api.GetDocument();
                        Asc.scope.p.forEach((item) => {
                            var oStyle = oDocument.GetStyle(item.type);
                            let oParagraph = Api.CreateParagraph();
                            oParagraph.SetStyle(oStyle);
                            oParagraph.AddText(item.content);
                            oDocument.InsertContent([oParagraph]);
                        });
                    })
                })
                .catch(error => {
                    console.error(error);
                })
                .finally(() => {
                    window.parent.parent.postMessage({"onlyOfficePlugin": "Loading"}, '*');
                });
            });
    });

    function sseRequest(question) {
        return new Promise((resolve, reject) => {

            fetch(
                "https://ai.azaas.com/api/v1/prediction/f240e13f-faf7-4cc1-8d57-7e6ff26580aa",
                {
                    method: "POST",
                    headers: {
                        "Content-Type": "application/json"
                    },
                    body: JSON.stringify({
                        "question": question,
                        "overrideConfig": {
                            "supabaseMetadataFilter": {
                                "supabase_0": {
                                    "rsdId": rsdId,
                                    "docType": "originalMaterial"
                                }
                            },
                            "memoryKey": rsdId,
                            "sessionId": rsdId
                        }
                    })
                }
            )
            // fetch("https://boa-admin-vnext.dev.azaas.online/api/ai/rsd/ai-prompt", {
            //     method: "POST",
            //     headers: {
            //         "Content-Type": "application/json",
            //         "Authorization": "Bearer " + token,
            //     },
            //     body: JSON.stringify({
            //         "RsdId": rsdId,
            //         "Question": question
            //     })
            // })
            .then(response => response.json())
            .then(result => resolve(result))
            .catch(error => reject(error));
        });

    }

    function generateText(text) {
        let lang = window.Asc.plugin.info.lang.substring(0, 2);
        return {
            en: text,
            [lang]: window.Asc.plugin.tr(text)
        }
    };

})(window, undefined);
