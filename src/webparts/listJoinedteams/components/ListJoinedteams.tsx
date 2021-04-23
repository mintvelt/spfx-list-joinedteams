import * as React from 'react';  
import styles from './ListJoinedteams.module.scss';  
import { IListJoinedteamsProps } from './IListJoinedteamsProps';  
import { IListJoinedteamsState } from './IListJoinedteamsState';  
import { escape } from '@microsoft/sp-lodash-subset';  
import { ITeam, ITeamCollection } from '../models/ITeam';  
import listJoinedTeamsService, { ListJoinedTeamsService } from '../services/ListJoinedTeamsService';  
import { List } from 'office-ui-fabric-react/lib/List';  
import { FocusZone, FocusZoneDirection } from 'office-ui-fabric-react/lib/FocusZone';  
import { TextField } from 'office-ui-fabric-react/lib/TextField';  
import { ITheme, mergeStyleSets, getTheme, getFocusStyle } from 'office-ui-fabric-react/lib/Styling';  
import { ShowIconsFieldLabel } from 'ListJoinedteamsWebPartStrings';
  
interface IGroupListClassObject {  
  itemCell: string;  
  itemImage: string;  
  itemContent: string;  
  itemName: string;  
  itemIndex: string;
  itemLink: string;  
  itemLinkTitle: string;  
  chevron: string;  
}  
  
const theme: ITheme = getTheme();  
const { palette, semanticColors, fonts } = theme;  
  
const classNames: IGroupListClassObject = mergeStyleSets({  
  itemCell: [  
    getFocusStyle(theme, { inset: -1 }),  
    {  
      minHeight: 54,  
      padding: 10,  
      boxSizing: 'border-box',  
      borderBottom: `1px solid ${semanticColors.bodyDivider}`,  
      display: 'flex',  
      selectors: {  
        '&:hover': { background: palette.neutralLight }  
      }  
    }  
  ],  
  itemImage: {  
    flexShrink: 0,
    float: "left",
    paddingRight: 10 
  },  
  itemContent: [
    fonts.small,
    {  
      marginLeft: 10,  
      overflow: 'hidden',  
      flexGrow: 1,
      color: '#8a8886'  
    }
  ],  
  itemName: [  
    fonts.medium,  
    {  
      whiteSpace: 'nowrap',  
      overflow: 'hidden',  
      textOverflow: 'ellipsis',
      color: '#323130' 
    }  
  ],  
  itemIndex: {  
    fontSize: fonts.small.fontSize,  
    color: palette.neutralTertiary,  
    marginBottom: 10  
  },  
  chevron: {  
    alignSelf: 'center',  
    marginLeft: 10,  
    color: palette.neutralTertiary,  
    fontSize: fonts.large.fontSize,  
    flexShrink: 0  
  },
  itemLinkTitle: {
    textDecoration: 'none',
    color: '#323130'
  },
  itemLink: {
    textDecoration: 'none',
    color: '#8a8886'
  }
});  
  
export default class ListJoinedteams extends React.Component<IListJoinedteamsProps, IListJoinedteamsState> {  
  private _originalItems: ITeam[] = []; 
  constructor(props: IListJoinedteamsProps) {  
    super(props);  
    this.state = {
      joinedTeams: [{ teamId: null, displayName: "Geen", description: "Geen teams gevonden" } ]
    };
  }  
  
  public componentDidMount(): void {  
    this._getJoinedTeams();  
  }  
  
  public render(): React.ReactElement<IListJoinedteamsProps> {  
    const { joinedTeams = [] } = this.state; 
    return (  
      <FocusZone direction={FocusZoneDirection.vertical}>  
        <h1>{this.props.description}</h1>  
        <List items={joinedTeams} onRenderCell={this._onRenderCell} />
      </FocusZone>  
    );  
  }  
  
  public _getJoinedTeams = (): void => {  
    listJoinedTeamsService.getJoinedTeams().then(result => {  
      this.setState({  
        joinedTeams: result  
      });  
      this._getTeamWebUrl(result);  
    });  
  }  
  
  public _getTeamWebUrl = (teams: any): void => {  
    teams.map(teamItem => (  
      listJoinedTeamsService.getTeamWebUrl(teamItem).then(teamUrl => {  
        if (teamUrl !== null) {  
          this.setState(prevState => ({  
            joinedTeams: prevState.joinedTeams.map(team => team.teamId === teamItem.teamId ? { ...team, webUrl: teamUrl.webUrl } : team)  
          }));  
        } 
      })  
    ));  
  }  
  
  private _onRenderCell(team: ITeam, index: number | undefined): JSX.Element {  
    return (
      <div className={classNames.itemCell} data-is-focusable={true}>  
        <div className={classNames.itemContent}>  
          <div className={classNames.itemImage}>
            <img src="https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/teams_32x1.svg" width="32" height="32" alt="Teams product icon" />
          </div>
          <div className={classNames.itemName}>  
            <a href={team.webUrl} target="_blank" className={classNames.itemLinkTitle}>{team.displayName}</a>  
          </div>  
          <div className={classNames.itemContent}>  
            <a href={team.webUrl} target="_blank" className={classNames.itemLink}>{team.description}</a>  
          </div>  
        </div>  
      </div>  
      );
  }  
  private _onRenderCellText(team: ITeam, index: number | undefined): JSX.Element {  
    return (
      <div className={classNames.itemCell} data-is-focusable={true}>  
        <div className={classNames.itemContent}>  
          <div className={classNames.itemName}>  
            <a href={team.webUrl} target="_blank" className={classNames.itemLinkTitle}>{team.displayName}</a>  
          </div>  
          <div className={classNames.itemContent}>  
            <a href={team.webUrl} target="_blank" className={classNames.itemLink}>{team.description}</a>  
          </div>  
        </div>  
      </div>  
      );
  } 
}  

