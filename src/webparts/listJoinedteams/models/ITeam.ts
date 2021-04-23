import { Guid } from '@microsoft/sp-core-library';  
  
export interface ITeam {  
    teamId: Guid;  
    displayName: string;  
    description: string;
    webUrl?: string;  
}  
  
export interface ITeamCollection {  
    value: ITeam[];  
}  