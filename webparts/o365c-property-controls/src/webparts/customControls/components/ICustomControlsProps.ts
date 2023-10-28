import { IHeaderProps } from "../interfaces/webpart.types";

export interface ICustomControlsProps {
  headerProps: IHeaderProps;
  listId: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
