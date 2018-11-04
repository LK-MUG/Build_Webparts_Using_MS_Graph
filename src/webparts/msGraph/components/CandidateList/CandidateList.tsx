import * as React from "react";
import styles from '../MsGraph.module.scss';
import { ICandidateListProps } from "./ICandidateListProps";
import { DetailsList, IColumn } from 'office-ui-fabric-react';
import { ISPListCandidate } from "../../services/responseTypes/ISPListCandidate";

const candidateListColumns: IColumn[] = [
    {
        fieldName: "Title",
        key:"Title",
        name:"Title",
        minWidth: 0,
        maxWidth: 1
    },
    {
        fieldName: "EMail",
        key:"EMail",
        name:"E-Mail",
        minWidth: 0,
        maxWidth: 1
    },
];

export class CandidateList extends React.Component<ICandidateListProps, {}> {

    public constructor(props: ICandidateListProps) {
        super(props);
        this.switchCandidate = this.switchCandidate.bind(this);
    }

    public switchCandidate(newCandiadte: ISPListCandidate, index: number): void {
        index++;
        if (this.props.selectedIndex != index) {
            this.props.switchCandidate(index);
        }
    }

    public render(): React.ReactElement<ICandidateListProps> {
        return (
            <div className={ styles.column }>
              <span className={ styles.title }>Candidates who applied</span>
              <DetailsList
                compact={true}
                items={this.props.users.map((user) => {
                    return {
                        Title: user.Title,
                        EMail: user.EMail,
                    };
                })}
                onActiveItemChanged={this.switchCandidate}
                columns={candidateListColumns}
              >
              </DetailsList>
            </div>
        );
    }
}