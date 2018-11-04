import * as React from 'react';
import styles from './MsGraph.module.scss';
import { ICandidateInternalInfoProps } from './ICandidateInternalInfoProps';
import { ICandidateInternalInfoState } from './ICandidateInternalInfoState';
import { IGraphCandidate } from '../services/responseTypes/IGraphCandidate';
import { ISPListCandidate, ISPListCandidateListItem } from '../services/responseTypes/ISPListCandidate';
import { CandidateList } from "./CandidateList/CandidateList";
import { CandidateInfo } from "./CandidateInfo/CandidateInfo";

export default class CandidateInternalInfo extends React.Component<ICandidateInternalInfoProps, ICandidateInternalInfoState> {
  
  constructor(props: ICandidateInternalInfoProps, state: ICandidateInternalInfoState) {
    super(props);
    
    // Initialize the state of the component
    this.state = {
      users: [],
      candidate: null,
      selectedId: -1,
    };

    // binding so that we know what this is
    this.getCandidateById = this.getCandidateById.bind(this);
  }

  private getListIDFromURL(): number {
    // retrieve the listItem ID from the Query Params, we don't care what the other pramas
    // in the url are, we are assuming there is only 1.
    let url = window.location.href.split('?');
    let idParam;
    if (url.length > 1) {
      try {
        idParam = parseInt(url[1].split('=')[1], 10);
      } catch (error) {
        console.error("Failed to find the Candidate ListItem ID from URL");        
      }
    }
    return idParam;
  }

  public componentDidMount(): void {
    const listId = this.getListIDFromURL();
    if (listId > 0) {
      this.getCandidateById(listId);
    } 
    this.getCandidateList();
  }

  private getCandidateList(): void {
    this.props.spListClient.getListOfCandidates()
    .then((availableCandidates: ISPListCandidateListItem[]) => {
      this.setState((prevState: ICandidateInternalInfoState) => {
        prevState.users = availableCandidates.map((currentItem) => {
          return currentItem.Author as ISPListCandidate;
        });
        return prevState;
      });
      console.info("Retrieved the List of all available candidates", availableCandidates);
    })
    .catch((error) => {
      console.error("Failed to retrieve a list of candidates, please check the property pane.");
    }); 
  }

  private getCandidateById(id: number): void {
    this.setState({selectedId: id});

    this.props.spListClient.getUserPrincipal(id)
    .then((userPrincipal: ISPListCandidate) => {
      console.info("Retrieved the userprincipal from the SharePoint List.", userPrincipal);
      return this.props.graphClient.findUser(userPrincipal.EMail);
    })
    .then((candidate: IGraphCandidate) => {
      console.info("Retrieved the candidate from MS_Graph", candidate);
      this.setState((prevState: ICandidateInternalInfoState) => {
        prevState.candidate = candidate;
        return prevState;
      });
    })
    .catch((error) => {
      console.error("Failed to retrive the candidate from MS_Graph using the SP List", error);
    });
  }

  public render(): React.ReactElement<ICandidateInternalInfoProps> {
    // The two sub-classes are just to make this one easier to read.
    return (
      <div className={ styles.graphConsumer }>
        <div className={ styles.container }>
          <div className={ styles.row }>
              <CandidateList
                selectedIndex={this.state.selectedId}
                switchCandidate={this.getCandidateById}
                users={this.state.users}
              />
              {
                this.state.candidate &&
                <CandidateInfo
                  candidate={this.state.candidate}
                  title={"Candidate Info"}
                />
              }
          </div>
        </div>
      </div>
    );
  }
}
