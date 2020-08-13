import { IQuestionProps } from "./QuestionComponent/IQuestionProps";
import { MessageBarType } from "office-ui-fabric-react/lib/MessageBar";

export interface ISurveyState {
    questions : IQuestionProps[];
    questionsCount : number;
    questionsCountLoaded : boolean;
    fillSurveyBtnClicked : boolean;
}