import { ICustomControlsWebPartProps } from "../../webparts/customControls/CustomControlsWebPart";
import { IHeaderProps } from "../../webparts/customControls/interfaces/webpart.types";

export default class WebpartMapper {
  public static mapHeader(
    properties: ICustomControlsWebPartProps
  ): IHeaderProps {
    return {
      title: properties.title,
      description: properties.description,
    } as IHeaderProps;
  }
}
