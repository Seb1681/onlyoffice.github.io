// an Chat plugin of AI
(function (window, undefined) {
    let messageHistory = null; // a reference to the message history DOM element
    let conversationHistory = null; // a list of all the messages in the conversation
    let messageInput = null;
    let typingIndicator = null;
    let lang = '';
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

        lang = window.Asc.plugin.info.lang.substring(0, 2);
        messageHistory = document.querySelector('.message-history');
        conversationHistory = [];
        typingIndicator = document.querySelector('.typing-indicator');
        const greet = "Hi I'm RSD pilot, your dedicated assistant for crafting precise and effective Requirement Specification Documents. How can I assist you today?";
        displayMessage(greet, 'ai-message');
    };

    // const start = async (connection) => {
	// 	try {
	// 		await connection.start();
	// 		console.log("SignalR Connected from RSD Pilot");
	// 	} catch (err) {
	// 		console.error(err);
	// 		// setTimeout(() => start(), 5000);
	// 	}
	// };

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
                        // {
                        //     id: 'summarize',
                        //     text: generateText('summarize'),
                        // },
                        // {
                        //     id: 'explain',
                        //     text: generateText('explain'),
                        // },
                        // {
                        //     id: 'translate',
                        //     text: generateText('translate'),
                        //     items: [
                        //         {
                        //             id: 'translate_to_en',
                        //             text: generateText('translate to English'),
                        //         },
                        //         {
                        //             id: 'translate_to_zh',
                        //             text: generateText('translate to Chinese'),
                        //         },
                        //         {
                        //             id: 'translate_to_fr',
                        //             text: generateText('translate to French'),
                        //         },
                        //         {
                        //             id: 'translate_to_de',
                        //             text: generateText('translate to German'),
                        //         },
                        //         {
                        //             id: 'translate_to_ru',
                        //             text: generateText('translate to Russian'),
                        //         },
                        //         {
                        //             id: 'translate_to_es',
                        //             text: generateText('translate to Spanish'),
                        //         }
                        //     ]
                        // },
                        // {
                        //     id: 'clear_history',
                        //     text: generateText('Clear chat history'),
                        // }
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

    window.Asc.plugin.attachContextMenuClickEvent('clear_history', function () {
        clearHistory();
    });

    const displayMessage = function (message, messageType) {
        message = message.replace(/^"|"$/g, ''); // remove surrounding quotes
        message = message.replace(/\\n/g, '\n'); // replace \n with newline characters

        // create a new message element
        const messageElement = document.createElement('div');
        messageElement.classList.add(messageType); // Add div class

        // split the message into lines and create a text node for each line
        const lines = message.split('\n');
        for (const line of lines) {
            const textNode = document.createTextNode(line);
            messageElement.appendChild(textNode);
            messageElement.appendChild(document.createElement('br'));
        }

        const containerElement = document.createElement('div');
        containerElement.classList.add(messageType + '-container'); // Add div class
        containerElement.appendChild(messageElement);
        // add the message element to the message history
        messageHistory.appendChild(containerElement);

        //  scroll to the bottom of the message history
        messageHistory.scrollTop = messageHistory.scrollHeight;
    };

    //summarize
    window.Asc.plugin.attachContextMenuClickEvent('summarize', function () {
        window.Asc.plugin.executeMethod('GetSelectedText', null, function (text) {
            conversationHistory.push({ role: 'user', content: prompts[lang]['summarize'] + text });
            sseRequest(conversationHistory)
                .then(result => {
                    console.log("success");
                    let currentDiv = null;
                    let currentMessage = null;
                    displaySSEMessage(result, currentDiv, currentMessage);
                })
                .catch(error => {
                console.log("error", error);
                });
        });
    });

    // explain 
    window.Asc.plugin.attachContextMenuClickEvent('explain', function () {
        window.Asc.plugin.executeMethod('GetSelectedText', null, function (text) {
            conversationHistory.push({ role: 'user', content: prompts[lang]['explain'] + text });
            sseRequest(conversationHistory)
                .then(result => {
                    console.log("success");
                    let currentDiv = null;
                    let currentMessage = null;
                    displaySSEMessage(result, currentDiv, currentMessage);
                })
                .catch(error => {
                console.log("error", error);
                });
            typingIndicator.style.display = 'none'; // hide the typing indicator
        });
    });

    const parseMarkdown = (markdownString) => {
        console.log('parsing markdown');
        const lines = markdownString.split('\n');
        console.log(lines);
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
            let prompt = (`Please generate the content for: "${text}". Please reply only the content in markdown format for this prompt.`);
            typingIndicator.innerHTML = 'Generating...';
            typingIndicator.style.display = 'block'; // display the typing indicator
            sseRequest(prompt)
                .then(result => {
                    Asc.scope.p = parseMarkdown(result);
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
                    typingIndicator.style.display = 'none'; // hide the typing indicator
                });
            });
    });

    // rephrase content in document
    window.Asc.plugin.attachContextMenuClickEvent('rephrase', function () {
        window.Asc.plugin.executeMethod('GetSelectedText', null, (text) => {
            window.parent.parent.postMessage({"onlyOfficePlugin": "Loading"}, '*');
            let prompt = (`Please rephrase this sentence: "${text}". Please reply only the content in markdown format for this prompt.`);
            typingIndicator.innerHTML = 'Rephrasing...';
            typingIndicator.style.display = 'block'; // display the typing indicator
            sseRequest(prompt)
                .then(result => {
                    Asc.scope.p = parseMarkdown(result);
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
                    typingIndicator.style.display = 'none'; // hide the typing indicator
                });
            });
    });
    // Make sure the DOM is fully loaded before querying the DOM elements
    document.addEventListener("DOMContentLoaded", function () {
        // get references to the DOM elements
        messageInput = document.querySelector('.message-input');
        const sendButton = document.querySelector('.send-button');
        typingIndicator = document.querySelector('.typing-indicator');

        // send a message when the user clicks the send button
        function sendMessage() {
            const message = messageInput.value;
            if (message.trim() !== '') {
                displayMessage(message, 'user-message');
                conversationHistory.push({ role: 'user', content: message });
                messageInput.value = '';
                typingIndicator.innerHTML = 'Thinking...';
                typingIndicator.style.display = 'block'; // display the typing indicator
                sseRequest(message)
                    .then(result => {
                        console.log("success");
                        displayMessage(result, 'ai-message');
                    })
                    .catch(error => {
                        console.log("error", error);
                    })
                    .finally(() => typingIndicator.style.display = 'none'); // hide the typing indicator
            }
        }

        sendButton.addEventListener('click', sendMessage);

        messageInput.addEventListener('keydown', function (event) {
            if (event.key === 'Enter') {
                event.preventDefault();  // prevent the default behavior of the Enter key
                if (event.shiftKey) {
                    // if the user pressed Shift+Enter, insert a newline character
                    messageInput.value += '\n';
                } else {
                    // if the user only pressed Enter, send the message
                    sendMessage();
                }
            }
        });
    });

    function clearHistory() {
        messageHistory.innerHTML = '';
        conversationHistory = [];
        messageInput.value = '';
    }

    function sseRequest(question) {
        return new Promise((resolve, reject) => {

            // fetch(
            //     "https://ai.azaas.com/api/v1/prediction/97bd8c9a-5f24-4bb2-8484-a0d3a3b8f041",
            //     {
            //         method: "POST",
            //         headers: {
            //             "Content-Type": "application/json"
            //         },
            //         body: JSON.stringify({
            //             "question": question,
            //             "overrideConfig": {
            //                 "supabaseMetadataFilter": {
            //                     "supabaseExistingIndex_0": {
            //                         "rsdId": rsdId,
            //                         "docType": "originalMaterial"
            //                     }
            //                 },
            //                 "memoryKey": {"bufferMemory_0": rsdId},
            //                 "inputKey": {"bufferMemory_0": rsdId},
            //             }
            //         })
            //     }
            // )
            fetch("https://boa-admin-vnext.dev.azaas.online/api/ai/rsd/ai-prompt", {
                method: "POST",
                headers: {
                    "Content-Type": "application/json",
                    "Authorization": "Bearer " + token,
                },
                body: JSON.stringify({
                    "RsdId": rsdId,
                    "Question": question
                })
            })
            .then(response => response.json())
            .then(result => resolve(result))
            .catch(error => reject(error));
        });

    }

    function displaySSEMessage(result, currentDiv, currentMessage) {
        // console.log("stream result: ", result);
        if (result) {
            // console.log("result.value of stream", result.value);
            conversationHistory.push({ role: 'assistant', content: currentMessage });
            return;
        }
        if (currentDiv === null) {
            currentDiv = document.createElement('div');
            currentDiv.classList.add('ai-message');
            messageHistory.appendChild(currentDiv);
        }
        if (currentMessage === null) {
            currentMessage = '';
        }
        const lines = result.value.split('\n');
        lines.forEach(line => {
            if (line.includes('data')) {
                const fragment = line.split(':')[1];
                currentMessage += fragment;
                if (fragment === '') {
                    currentDiv.appendChild(document.createElement('br'));
                } else {
                    currentDiv.appendChild(document.createTextNode(fragment));
                }
            }
        });

        // recursively call processResult() to continue reading data from the stream
        displaySSEMessage(result, currentDiv, currentMessage);
    }

    function generateText(text) {
        let lang = window.Asc.plugin.info.lang.substring(0, 2);
        return {
            en: text,
            [lang]: window.Asc.plugin.tr(text)
        }
    };

})(window, undefined);
