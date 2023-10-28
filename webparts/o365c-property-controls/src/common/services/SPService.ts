import { SPFI } from "@pnp/sp";
import { LogHelper } from "../helpers/LogHelper";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Web } from "@pnp/sp/webs";

class SPService {
  private static _sp: SPFI;

  public static Init(sp: SPFI): void {
    this._sp = sp;
    LogHelper.info("SPService", "constructor", "PnP SP context initialised");
  }
  public static getLists = async (siteUrl: string): Promise<any> => {
    try {
      const web = Web([this._sp.web, siteUrl]);
      // gets the web info
      //const webInfo = await web();
      const lists = (await web.lists()).filter(
        (l) => l.BaseTemplate === 100 && l.BaseType === 0 && !l.Hidden
      );
      console.log("Lists:", lists);
      return lists;
    } catch (err) {
      LogHelper.error("SPService", "getLists", err);
      return null;
    }
  };
}
export default SPService;
