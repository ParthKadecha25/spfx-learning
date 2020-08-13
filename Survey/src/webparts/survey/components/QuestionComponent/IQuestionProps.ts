import { QuestionTypesEnum } from "../QuestionTypesEnum";
import { SPHttpClient } from "@microsoft/sp-http";
import { QuestionModeEnum } from "../QuestionModeEnum";

export interface IQuestionProps{
    listName: string;
    spHttpClient: SPHttpClient;
    siteUrl: string;
    questionMode : QuestionModeEnum
    questionType? : QuestionTypesEnum;
    questionNo? : number;
    questionID? : number;
    required? : boolean;
    questionText? : string;
    answerText? : string;
    longTextAnswer? : boolean;
    choices? : string[];
    ratingLevel? : number;
    onDelete? : any;
    onUpdate? : any;
    adminMode : boolean;
}