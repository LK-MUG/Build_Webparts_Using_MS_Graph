import { ISPListCandidate } from "../../services/responseTypes/ISPListCandidate";

export interface ICandidateListProps {
    users: ISPListCandidate[];
    switchCandidate: (listItemIdx: number) => void; 
    selectedIndex: number;
}