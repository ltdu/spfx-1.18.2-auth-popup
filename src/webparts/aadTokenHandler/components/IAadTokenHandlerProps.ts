import { BaseComponentContext } from "@microsoft/sp-component-base";

export interface IAadTokenHandlerProps {
  context: BaseComponentContext;
  redirectionRequired: boolean;
  redirectionUrl: string;
  popupRequired: boolean;
  popup: () => void;
  invoke: () => Promise<void>;
  log: string[]
}
