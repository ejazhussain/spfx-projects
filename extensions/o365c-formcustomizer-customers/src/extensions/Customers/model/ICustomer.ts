import { ITermData } from "./ITermData";

export interface ICustomer {
    ID: string
    Title: string;
    Email: string;
    Address?: string;
    Interests?: string[];
    ProjectsId?: string[];
    Projects?: any[];
    CustomerLocations?: ITermData[];
}