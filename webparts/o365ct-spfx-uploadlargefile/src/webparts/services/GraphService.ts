import { MSGraphClientV3 } from "@microsoft/sp-http";
import FileHelper from "../helpers/FileHelper";



export interface IGraphService {
  GetUserProfile(): Promise<void>;
  GetEvents(): Promise<any>;
}

export default class GraphService {

  private static _instance: GraphService = null;
  private static _graphClient: MSGraphClientV3;

  private constructor(graphClient: MSGraphClientV3) {
    GraphService._graphClient = graphClient;
  }
  // Call this from any web part whose components will require
  // use of this service.
  public static init(msGraphClient: MSGraphClientV3): void {
    if (GraphService._instance === null) {
      GraphService._instance = new GraphService(msGraphClient);
    }
  }

  public static async GetUserProfile(): Promise<any> {
    try {

      let userResponse: any = await this._graphClient.api("/me").get();
      let photoResponse: any = await this._graphClient.api("/me/photo/$value").get();

      let user = {
        name: userResponse.displayName,

        email: userResponse.mail,
        phone: userResponse.businessPhones[0],
        photo: window.URL.createObjectURL(photoResponse)
      };

      return user;

    }
    catch (error) {
      console.log("Error in GetUserProfile:", error);
      return null;
    }
  }


  /**
   * Upload files [less than 4MB in size] to OneDrive
   * @param file 
   * @returns Uploaded file
   */
  public static async UploadSmallFile(file: File): Promise<any> {
    try {
      const fileContent = await FileHelper.readFileContent(file);
      let uploadedFile: any = await this._graphClient.api(`/me/drive/root:/${file.name}:/content`).put(fileContent);

      return uploadedFile;

    }
    catch (error) {
      console.log("Error in UploadSmallFile:", error);
      return null;
    }
  } 

}
