/**
 *
 * (c) Copyright Ascensio System SIA 2020
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 */
(function(window, undefined){
    
    window.Asc.plugin.init = async function()
    {
		const uuid = uuidv4();
		const payload = {
			onlyOfficePlugin: 'OverwriteContent',
			pluginId: uuid
		}

		const connection = new signalR.HubConnectionBuilder()
			.withUrl("https://boa-admin-vnext.dev.azaas.online/signalr-hubs/onlyOffice?pluginId=" + uuid, {
				skipNegotiation: true,
				transport: signalR.HttpTransportType.WebSockets
			  })
			// .configureLogging(signalR.LogLevel.Information)
			.build();

		connection.on("ReceiveMessage", (user, message) => {
			console.log('Message on overwrite: ' + message);
			overwriteContent(message);

			connection.stop().then(() => console.log('Overwrite-content connection successfully closed.')).catch(err => console.error('Error while closing the connection: ', err));
		});
        console.log('Overwrite Plugin Initiated')
		await connection.start();
		console.log("SignalR Connected on overwrite content");
		window.parent.parent.postMessage(payload, '*');
    };

    window.Asc.plugin.button = function(id)
    {
		this.executeCommand("close", "");
    };
	
	const parseMarkdown = (markdownString) => {
        const lines = markdownString.split('\\n');
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
			else {
				return { type: 'Normal', content: line };
			}
    	});
    	return categorized;
    }
	
	const overwriteContent = (msg) => {
		if (msg) {
			console.log('Overwrite Received: ');
			if (msg) {
				Asc.scope.p = parseMarkdown(msg.slice(1, -1));
				Asc.plugin.callCommand(() => {
					var oDocument = Api.GetDocument();
					oDocument.RemoveAllElements();
					Asc.scope.p.forEach((item, index) => {
						var oStyle = oDocument.GetStyle(item.type);
						var oParagraph = Api.CreateParagraph();
						oParagraph.SetStyle(oStyle);
						oParagraph.AddText(item.content);
						oDocument.AddElement(index, oParagraph);
					});
				})
			}
		}
	}

	const start = async (connection) => {
		try {
			await connection.start();
			console.log("SignalR Connected on overwrite content");
		} catch (err) {
			console.error(err);
			// setTimeout(() => start(), 5000);
		}
	};
})(window, undefined);
