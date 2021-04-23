import { MSGraphClient } from "@microsoft/sp-http";  
import { WebPartContext } from "@microsoft/sp-webpart-base";  
import { ITeam, ITeamCollection } from '../models/ITeam';  
  
export class ListJoinedTeamsService {  
  public context: WebPartContext;  
  
  public setUp(context: WebPartContext): void {  
    this.context = context;  
  }  
  
  public getJoinedTeams(): Promise<ITeam[]> {  
    return new Promise<ITeam[]>((resolve, reject) => {  
      try {  
        // Prepare the output array    
        var teams: Array<ITeam> = new Array<ITeam>();  
  
        this.context.msGraphClientFactory  
          .getClient()  
          .then((client: MSGraphClient) => {  
            client  
              .api("/me/joinedTeams")  
              .select('id,displayName,description')  
              .get((error: any, teamColl: ITeamCollection, rawResponse: any) => {  
                // Map the response to the output array    
                teamColl.value.map((item: any) => {  
                  teams.push({  
                    teamId: item.id,  
                    displayName: item.displayName,
                    description: item.description  
                  });  
                });  
                resolve(teams);  
              });  
          });  
      } catch (error) {  
        console.error(error);  
      }  
    });  
  }  
  
  public getTeamWebUrl(teams: ITeam): Promise<any> {  
    return new Promise<any>((resolve, reject) => {  
      try {  
        this.context.msGraphClientFactory  
          .getClient()  
          .then((client: MSGraphClient) => {  
            client  
              .api(`/teams/${teams.teamId}`)  
              .select('webUrl')  
              .get((error: any, team: any, rawResponse: any) => {  
                resolve(team);  
              });  
          });  
      } catch (error) {  
        console.error(error);  
      }  
    });  
  }  
}  
  
const listJoinedTeamsService = new ListJoinedTeamsService();  
export default listJoinedTeamsService;  