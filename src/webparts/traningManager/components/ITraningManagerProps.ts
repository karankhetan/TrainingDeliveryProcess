import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ITraningManagerProps {
  description: string;
  getlistItem:()=>Promise<any[]>;
  DeleteListItem:(id:any)=>Promise<any>;
  context:WebPartContext;
}
