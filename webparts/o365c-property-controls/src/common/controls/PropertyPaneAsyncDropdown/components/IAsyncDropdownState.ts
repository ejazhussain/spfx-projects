import { IDropdownOption } from "@fluentui/react";

export interface IAsyncDropdownState {
  loading: boolean;
  options: IDropdownOption[];
  error: string | undefined;
}
