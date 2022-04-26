/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

// const sso = require("office-addin-sso");

import dialogFallback from "./fallbackAuthTaskpane";

async function getGraphToken() {
  if (Office.context.requirements.isSetSupported("IdentityAPI", "1.3")) {
    console.log("Inside isSetSupported!!!!!!!!!");
    try {
      let userTokenEncoded = await OfficeRuntime.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: true,
      }); //{allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true}
      console.log(userTokenEncoded);
      // let exchangeResponse = await sso.getGraphToken(userTokenEncoded);
      // console.log("exchangeResponse ---->>>>>" + exchangeResponse);
      // if (exchangeResponse.claims) {
      //   let mfaBootstrapToken = await OfficeRuntime.auth.getAccessToken({ authChallenge: exchangeResponse.claims });
      //   console.log("mfaBootstrapToken ---->>>>>" + mfaBootstrapToken);
      //   exchangeResponse = sso.getGraphToken(mfaBootstrapToken);
      // }
      //
      // if (exchangeResponse.error) {
      //   console.error(exchangeResponse);
      // } else {
      //   return exchangeResponse.access_token;
      // }
    } catch (exception) {
      console.error(exception);
      dialogFallback();
    }
  }
}

Office.initialize = function (reason) {
  // ssoAuthHelper.getGraphData();
  getGraphToken()
    .then((tkn) => {
      console.log(tkn);
      // fetch("https://graph.microsoft.com/v1.0/me", {
      //   method: "get",
      //   headers: new Headers({
      //     Authorization: "Bearer " + tkn,
      //     "Content-Type": "application/json",
      //   }),
      // }).then((response) => {
      //   console.log(response);
      //   item.body.prependAsync("Token - " + tkn + "User - " + response, {
      //     coercionType: Office.CoercionType.Html,
      //   });
      // });
    })
    .catch((err) => {
      console.error(err);
    });
};

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  // getUserData().then((res) => {
  //   console.log(res);
  // });
}
