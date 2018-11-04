import { ICandidateGraphService } from "../services/ICandidateGraphServices";
import { ICandidateListService } from "../services/ICandidateListService";

/**
 * Dependencies/Properties of the Appliaction.
 * We use interfaces as a technique to decouple our React Web App
 * from SharePoint. 
 */
export interface ICandidateInternalInfoProps {
  graphClient: ICandidateGraphService;
  spListClient: ICandidateListService;
}
