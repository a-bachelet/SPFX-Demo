import { IUserService } from "./IUserService";
import { User } from "@microsoft/microsoft-graph-types";
import { MSGraphClient } from "@microsoft/sp-http";

export class UserService implements IUserService {

  private static _instance: UserService;

  private _msGraphClient: MSGraphClient;

  private constructor(msGraphClient: MSGraphClient) {
    this._msGraphClient = msGraphClient;
  }

  public static getInstance(msGraphClient: MSGraphClient): UserService {
    if ( this._instance == null ) {
      this._instance = new UserService(msGraphClient);
    }

    return this._instance;
  }

  public getProfile(): Promise<User> {
    return new Promise<User>((resolve, reject) => {

      this._msGraphClient
        .api('/me')
        .version('v1.0')
        .get((error: any, response: User, rawResponse: any) => {
          resolve(response);
          reject(error);
        });
    });
  }

}
