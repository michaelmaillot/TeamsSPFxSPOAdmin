import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IAdminSpoProps {
  description: string;
  hasTeamsContext: boolean;
  context: WebPartContext;
}
