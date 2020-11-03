/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global document, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  var item = Office.context.mailbox.item;
  // var itembody;
  // Office.context.mailbox.item.Body.getTypeAsync(function(asyncResult) {
  //   if (asyncResult.status === "failed") {
  //     console.log("Action failed with error: " + asyncResult.error.message);
  //   } else {
  //     itembody = asyncResult.value;
  //     console.log("Body type: " + asyncResult.value);
  //   }
  // });

  document.getElementById("item-sender").innerHTML =
    "<b>Sender:</b> <br/>" + item.sender.displayName + "(" + item.sender.emailAddress + ")";
  document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
  // document.getElementById("item-body").innerHTML = "<b>Body:</b> <br/>" + itembody;
}
