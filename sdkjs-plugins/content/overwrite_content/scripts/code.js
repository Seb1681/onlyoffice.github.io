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
    
    window.Asc.plugin.init = function()
    {
		const uuid = uuid.v4();
		console.log(uuid, 'uuid');
		const payload = {
			onlyOfficePlugin: 'OverwriteContent',
			pluginId: uuid
		}
		window.parent.parent.postMessage(payload, '*');

		const connection = new signalR.HubConnectionBuilder()
			.withUrl("http://localhost:44301/signalr-hubs/onlyOffice", {
				skipNegotiation: true,
				transport: signalR.HttpTransportType.WebSockets
			  })
			.configureLogging(signalR.LogLevel.Information)
			.build();

		connection.on("ReceiveMessage", (user, message) => {
			console.log('User on Plugin: ' + user);
			console.log('Message on Plugin: ' + message);
		});

		start(connection).then(() => {console.log('SignalR Connected')});
    };

    window.Asc.plugin.button = function(id)
    {
		this.executeCommand("close", "");
    };
	
	window.Asc.plugin.event_onDocumentContentReady = function()
	{
		// const overwriteContent = (event) => {
		// 	const msg = event.data;
		// 	if (msg && typeof msg === 'object' && msg.action && msg.action == 'overwriteContent') {
		// 		console.log('Overwrite Received: ');
		// 		console.log(msg);
		// 		if (msg.content) {
		// 			Asc.scope.msgContent = msg.content;
		// 			Asc.plugin.callCommand(() => {
		// 				var oDocument = Api.GetDocument();
		// 				oDocument.RemoveAllElements();
		// 				var oParagraph = Api.CreateParagraph();
		// 				oParagraph.AddText(Asc.scope.msgContent);
		// 				oDocument.AddElement(0, oParagraph);
		// 			})
		// 		}
		// 		return true;
		// 	}
		// 	return false;
		// }
		// window.addEventListener('message', event => {
		// 	// IMPORTANT: check the origin of the data!
		// 	// if (event.origin === 'https://your-first-site.example') {
		// 		// The data was sent from your site.
		// 		// Data sent with postMessage is stored in event.data:
		// 	// }

		// 	// Remove event listener after executing 1 time;
		// 	if (overwriteContent(event)) {
		// 		window.removeEventListener('message', event => {
		// 			overwriteContent(event);
		// 		});
		// 	}
		// });
	};

	async function start(connection) {
		try {
			await connection.start();
			console.log("SignalR Connected.");
		} catch (err) {
			console.log('Error from Plugin')
			console.error(err);
			// setTimeout(() => start(), 5000);
		}
	};
})(window, undefined);
