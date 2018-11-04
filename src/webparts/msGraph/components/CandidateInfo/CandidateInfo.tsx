import * as React from "react";
import { ICandidateInfoProps } from "./ICandidateInfoProps";
import styles from '../MsGraph.module.scss';
import { DetailsList, CheckboxVisibility, IColumn } from 'office-ui-fabric-react';
import { IGraphCandidate } from "../../services/responseTypes/IGraphCandidate";

const candidateInfoColumns: IColumn[] = [
    {
        fieldName: "displayName",
        key:"displayName",
        name:"Name",
        minWidth: 0,
        maxWidth: 1
    },
    {
        fieldName: "mail",
        key:"mail",
        name:"E-Mail",
        minWidth: 0,
    },
    {
        fieldName: "mobilePhone",
        key:"mobilePhone",
        name:"Mobile Phone",
        minWidth: 0,
    },
    {
        fieldName: "businessPhone",
        key:"businessPhone",
        name:"Business Phone",
        minWidth: 0,
    },
];

export class CandidateInfo extends React.Component<ICandidateInfoProps, any> {
    public constructor(props: ICandidateInfoProps) {
        super(props);
    }

    public render(): React.ReactElement<ICandidateInfoProps> {
        return (
            <div className={ styles.column }>
              <span className={ styles.title }>{this.props.title}</span>
                <p className={ styles.form }>
                    <DetailsList
                        compact={false}
                        checkboxVisibility={CheckboxVisibility.hidden}
                        items={[this.props.candidate].map((candidate: IGraphCandidate) => {
                            return {
                                displayName: candidate.displayName,
                                mail: candidate.mail,
                                mobilePhone: candidate.mobilePhone,
                                businessPhone: candidate.businessPhones[0]
                            };
                        })}
                        columns={candidateInfoColumns}
                    />
                </p>
            </div>
        );
    }
}