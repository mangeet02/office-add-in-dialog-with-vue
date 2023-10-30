Office.initialize = function () {};

try {
  Office.actions.associate("main", main);
} catch (e) {
  console.log("Impossibile chiamare 'Office.actions.associate'", e);
}

let clickEvent;

function main(event) {
  clickEvent = event;
  return;
}

function eventHandler(arg) {
  console.log("Ricevuto evento dalla finestra di dialogo di avviso.");
  console.table(arg);
  clickEvent.completed();
}

async function messageHandler(arg) {
  console.log(`Ricevuto messaggio "${arg.message}" dalla finestra di dialogo di avviso.`);

  let itemId = Office.context.mailbox.item.itemId;
  let folderId = "JunkEmail";
  Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, async (result) => {
    const token = result.value;
    let res = await moveEmail(itemId, folderId, token);
    console.log(`Risposta API dalla richiesta di spostamento dell'elemento in ${folderId}:`);
    console.table(res);
    clickEvent.completed();
  });
}

async function moveEmail(itemId, folderId, token) {
  console.log(`Spostamento email con ID elemento ${itemId} in ${folderId}.`);
  const itemUrl = Office.context.mailbox.restUrl + "/v2.0/me/messages/";
  const restItemId = Office.context.mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0);
  const junkUrl = itemUrl + restItemId + "/move";

  let res = await fetch(junkUrl, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json;charset=utf-8',
      'Authorization': `Bearer ${token}`
    },
    body: JSON.stringify({
      "DestinationId": folderId
    })
  });
  return res;
}
