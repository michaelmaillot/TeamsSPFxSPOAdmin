import { WebPartContext } from "@microsoft/sp-webpart-base";

import { spfi, SPFI, SPFx } from "@pnp/sp";
import { LogLevel, PnPLogging } from "@pnp/logging";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";

import { MSAL } from "@pnp/msaljsclient";
import { Configuration, AuthenticationParameters } from "msal";
import Constants from "Constants";

let _sp: SPFI = null;
let _spAdmin: SPFI = null;

const configuration: Configuration = {
  auth: {
    authority: "https://login.microsoftonline.com/1fdd85e0-9a94-4593-8ab0-5ad1b834475f/",
    clientId: "d355e6e0-0aec-4b1f-b825-7040f81bc226"
  }
};

const authParams: AuthenticationParameters = {
  scopes: [Constants.TENANT_ADMIN_URL + "/.default"]
};

export const getSP = (context?: WebPartContext): SPFI => {
  if (_sp === null && context && context !== null) {
    _sp = spfi().using(SPFx(context), PnPLogging(LogLevel.Warning));
  }

  return _sp;
};

export const getSPAdmin = (context?: WebPartContext): SPFI => {
  if (_spAdmin === null && context && context !== null) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    _spAdmin = spfi(Constants.TENANT_ADMIN_URL).using(SPFx(context), MSAL(configuration as any, authParams), PnPLogging(LogLevel.Warning));
  }

  return _spAdmin;
}