import { QuestionTypesEnum } from "../QuestionTypesEnum";
import { QuestionModeEnum } from "../QuestionModeEnum";

export interface IQuestionState{
    questionMode : QuestionModeEnum;
    questionType : QuestionTypesEnum;
    questionText : string;
    required : boolean;
    answerText : string;
    choices : string[];
    currentChoice : string;
    ratingLevel : number;
    longTextAnswer : boolean;
    hideDeleteConfirmation : boolean;
}