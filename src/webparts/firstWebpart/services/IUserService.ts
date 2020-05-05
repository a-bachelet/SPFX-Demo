import { User } from '@microsoft/microsoft-graph-types';

export interface IUserService {

  getProfile(): Promise<User>;

}
