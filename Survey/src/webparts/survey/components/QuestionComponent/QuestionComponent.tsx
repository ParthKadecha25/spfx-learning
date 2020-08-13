import * as React from 'react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { IQuestionProps } from './IQuestionProps';
import { QuestionTypesEnum } from '../QuestionTypesEnum';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { ChoiceGroup } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { IQuestionState } from './IQuestionState';
import styles from './Question.module.scss'
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { ActionButton, IconButton } from 'office-ui-fabric-react/lib/Button';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { Dropdown } from 'office-ui-fabric-react/lib/Dropdown';
import { Rating, RatingSize } from 'office-ui-fabric-react/lib/Rating';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { QuestionModeEnum } from '../QuestionModeEnum';
import { DatePicker} from 'office-ui-fabric-react/lib/DatePicker';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';

export class QuestionComponent extends React.Component<IQuestionProps, IQuestionState>{

    private listItemEntityTypeName: string = undefined;

    constructor(prop) {

        // Initializing state value
        super(prop);
        this.state = this.setQuestionDetails();

        // Binding methods
        this.setQuestionDetails = this.setQuestionDetails.bind(this);
        this.addChoice = this.addChoice.bind(this);
        this.getChoices = this.getChoices.bind(this);

        this.saveQuestion = this.saveQuestion.bind(this);
        this.createQuestion = this.createQuestion.bind(this);
        this.setEditQuestionMode = this.setEditQuestionMode.bind(this);
        this.updateQuestion = this.updateQuestion.bind(this);

        this.deleteQuestionConfirmation = this.deleteQuestionConfirmation.bind(this);
        this.deleteQuestion = this.deleteQuestion.bind(this);
        this.closeDeleteQuestionConfirmation = this.closeDeleteQuestionConfirmation.bind(this);

        this.getListItemEtagValue = this.getListItemEtagValue.bind(this);
        this.getListItemEntityTypeName = this.getListItemEntityTypeName.bind(this);
    }

    render(): React.ReactElement<IQuestionProps> {

        let deleteConfirmationDialog = 
        <Dialog
            hidden={this.state.hideDeleteConfirmation}
            onDismiss={this.closeDeleteQuestionConfirmation}
            dialogContentProps={{
            type: DialogType.normal,
            title: 'Delete Question?',
            subText: 'Are you sure want to delete selected Question?'
            }}
            modalProps={{
            isBlocking: true,
            containerClassName: 'ms-dialogMainOverride'
            }}
        >
        <DialogFooter>
          <PrimaryButton onClick={this.deleteQuestion.bind(this, this.props.questionID)} text="Delete" />
          <DefaultButton onClick={this.closeDeleteQuestionConfirmation} text="Cancel" />
        </DialogFooter>
      </Dialog>

        let element;
        if (this.state.questionType == QuestionTypesEnum.Text) {
            element = <div className={[styles.row, "ms-Grid-row"].join(' ')}>
                <div className="ms-Grid-col ms-sm12">
                    {(this.state.questionMode == QuestionModeEnum.Create || this.state.questionMode == QuestionModeEnum.Edit) &&
                        <TextField type="text" value={this.state.answerText} placeholder="Sample Answer content. (Will be dislayed as hint)" readOnly={false} multiline={this.state.longTextAnswer} rows={4} onChanged={(value) => { this.setState({ answerText: value }) }}></TextField>
                    }
                    {this.state.questionMode == QuestionModeEnum.View &&
                        <TextField type="text" placeholder={this.state.answerText} readOnly={false} multiline={this.state.longTextAnswer} rows={4}></TextField>
                    }
                </div>
            </div>
        }
        else if (this.state.questionType == QuestionTypesEnum.Choice) {
            if (this.state.questionMode == QuestionModeEnum.Create || this.state.questionMode == QuestionModeEnum.Edit) {
                let choiceEntries;
                if (this.state.choices != null && this.state.choices.length > 0) {
                    choiceEntries = this.state.choices.map((choice, index) =>
                        <div className={[styles.row, "ms-Grid-row"].join(' ')}>
                            <div className="ms-Grid-col ms-sm1">
                                <Icon iconName="RadioBtnOff" className={styles.radioIconStyle} />
                            </div>
                            <div className="ms-Grid-col ms-sm9">
                                <TextField type="text" value={choice} readOnly={true}></TextField>
                            </div>
                            <div className="ms-Grid-col ms-sm2">
                                <IconButton iconProps={{ iconName: 'delete' }} title="Remove" ariaLabel="Remove" onClick={this.removeChoice.bind(this, this.state.choices[index])} />
                            </div>
                        </div>
                    );
                }
                const placeholder = "Choice " + (this.state.choices.length + 1);
                element =
                    <div>
                        {choiceEntries}
                        <div className={[styles.row, "ms-Grid-row"].join(' ')}>
                            <div className="ms-Grid-col ms-sm1">
                                <Icon iconName="RadioBtnOff" className={styles.radioIconStyle} />
                            </div>
                            <div className="ms-Grid-col ms-sm8">
                                <TextField type="text" value={this.state.currentChoice} placeholder={placeholder} onChanged={(value) => { this.setState({ currentChoice: value }) }} ></TextField>
                            </div>
                            <div className="ms-Grid-col ms-sm3">
                                <ActionButton text="Add Choice" iconProps={{ iconName: "add" }}
                                    className={styles.addChoiceButton}
                                    onClick={this.addChoice.bind(this, this.state.currentChoice)}
                                />
                            </div>
                        </div>
                    </div>
            }
            else if (this.state.questionMode == QuestionModeEnum.View) {
                element = <div className={[styles.row, "ms-Grid-row"].join(' ')}>
                    <div className="ms-Grid-col ms-sm12">
                        <ChoiceGroup options={this.getChoices(this.state.choices)} />
                    </div>
                </div>
            }
        }
        else if (this.state.questionType == QuestionTypesEnum.Rating) {
            const ratingLevel = <div className={[styles.row, "ms-Grid-row"].join(' ')}>
                <div className="ms-Grid-col ms-sm12">
                    {(this.state.questionMode == QuestionModeEnum.Create || this.state.questionMode == QuestionModeEnum.Edit) &&
                        <Rating min={1} max={this.state.ratingLevel} rating={5} size={RatingSize.Large} readOnly={true} />
                    }
                    {this.state.questionMode == QuestionModeEnum.View &&
                        <Rating min={1} max={this.state.ratingLevel} rating={this.state.ratingLevel} size={RatingSize.Large} readOnly={false} />
                    }
                </div>
            </div>

            let ratingLevelDropdown;
            if (this.state.questionMode == QuestionModeEnum.Create || this.state.questionMode == QuestionModeEnum.Edit) {
                ratingLevelDropdown = <div className={[styles.row, "ms-Grid-row"].join(' ')}>
                    <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg3">
                        <Dropdown
                            label="Max Level/Rating"
                            selectedKey={this.state.ratingLevel}
                            options={[
                                { key: 1, text: '1' },
                                { key: 2, text: '2' },
                                { key: 3, text: '3' },
                                { key: 4, text: '4' },
                                { key: 5, text: '5' }
                            ]}
                            onChanged={(option, value) => { this.setState({ ratingLevel: parseInt(option.key.toString()) }) }}
                        />
                    </div>
                </div>
            }
            element =
                <div>
                    {ratingLevel}
                    {ratingLevelDropdown}
                </div>
        }
        else if (this.state.questionType == QuestionTypesEnum.Date) {
            element =
                <div className={[styles.row, "ms-Grid-row"].join(' ')}>
                    <div className="ms-Grid-col ms-sm12 ms-md6">
                        {(this.state.questionMode == QuestionModeEnum.Create || this.state.questionMode == QuestionModeEnum.Edit) &&
                            <DatePicker showWeekNumbers={false} showMonthPickerAsOverlay={true} placeholder="Select a date..." ariaLabel="Select a date" disabled={true} />
                        }
                        {(this.state.questionMode == QuestionModeEnum.View) &&
                            <DatePicker showWeekNumbers={false} showMonthPickerAsOverlay={true} placeholder="Select a date..." ariaLabel="Select a date"/>
                        }
                    </div>
                </div>
        }

        if (this.state.questionMode == QuestionModeEnum.Create || this.state.questionMode == QuestionModeEnum.Edit) {
            return (
                <div className={styles.questionBox} dir="ltr">
                    <div className={[styles.row, "ms-Grid-row"].join(' ')}>
                        <ChoiceGroup
                            label="Question Type : "
                            onChange={(ev, option) => { this.setState({ questionType: option.key as QuestionTypesEnum.Choice }) }}
                            options={[
                                { key: QuestionTypesEnum.Text, text: 'Text', iconProps: { iconName: 'textfield' } },
                                { key: QuestionTypesEnum.Choice, text: 'Choice', iconProps: { iconName: 'RadioBtnOn' } },
                                { key: QuestionTypesEnum.Rating, text: 'Rating', iconProps: { iconName: 'Like' } },
                                { key: QuestionTypesEnum.Date, text: 'Date', iconProps: { iconName: 'DateTime' } }
                            ]}
                            selectedKey={this.state.questionType}
                        />
                    </div>
                    <div className={[styles.row, "ms-Grid-row"].join(' ')}>
                        <TextField label="Question :" placeholder="Type your question here" value={this.state.questionText} required={true} onChanged={(value) => { this.setState({ questionText: value }) }} />
                    </div>
                    <div className={[styles.row, "ms-Grid-row"].join(' ')}>
                        <Label required={this.state.required}>Answer : </Label>
                        {element}
                    </div>
                    <div className={[styles.row, "ms-Grid-row"].join(' ')}>
                        <hr />
                    </div>
                    <div className={[styles.row, "ms-Grid-row"].join(' ')}>
                        <div className={[styles.alignLeft, "ms-Grid-col ms-sm8"].join(' ')}>
                            {this.state.questionType == QuestionTypesEnum.Text &&
                                <Toggle onText="Long Answer" offText="Long Answer" checked={this.state.longTextAnswer}
                                    onChanged={(value) => { this.setState({ longTextAnswer: value }) }} />
                            }
                        </div>
                        <div className={[styles.alignRight, "ms-Grid-col ms-sm4"].join(' ')}>
                            <Toggle onText="Required" offText="Required" checked={this.state.required}
                                onChanged={(value) => { this.setState({ required: value }) }} />
                        </div>
                    </div>
                    <div className={[styles.row, "ms-Grid-row"].join(' ')}>
                        <div className="ms-Grid-col ms-sm12">
                            <DefaultButton primary={true} text="Save" onClick={this.saveQuestion} />
                            <DefaultButton primary={false} text="Cancel" onClick={() => this.setState(this.setQuestionDetails)} />
                        </div>
                    </div>
                </div>
            )
        }
        else {
            return (
                <div className={styles.questionBoxViewModeParent}>
                    <div className={styles.questionBoxViewMode}>
                        <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-sm10">
                                <Label style={{ fontWeight: "bold" }} required={this.state.required}>
                                    {this.props.questionNo}. {this.state.questionText}
                                </Label>
                            </div>
                            {this.props.adminMode && 
                                (
                                    <div className="ms-Grid-col ms-sm2">
                                        <IconButton title="Delete" iconProps={{ iconName: "Delete" }} onClick={this.deleteQuestionConfirmation.bind(this, this.props.questionID)} ></IconButton>
                                        &nbsp;
                                        <IconButton title="Edit" iconProps={{ iconName: "Edit" }} onClick={this.setEditQuestionMode.bind(this, this.props.questionID)} ></IconButton>
                                    </div>
                                )
                            }
                        </div>
                        <div className="ms-Grid-row">
                            {element}
                        </div>
                        {deleteConfirmationDialog}
                    </div>
                </div>
            )
        }
    }

    addChoice(choice: string) {
        if (choice != "") {
            let newChoices = this.state.choices;
            newChoices.push(choice);
            this.setState({
                choices: newChoices,
                currentChoice: "",
            });
        }
    }

    removeChoice(choiceToRemove: string) {
        if (choiceToRemove != "") {
            let newChoices = [];
            this.state.choices.map((choice) => {
                if (choice != choiceToRemove) {
                    newChoices.push(choice);
                }
            });
            this.setState({
                choices: newChoices,
            });
        }
    }

    getChoices(choices: string[]) {
        let result = [];
        choices.forEach((element, index) => {
            let choice = {
                key: index.toString(),
                text: element
            }
            result.push(choice);
        });
        return result;
    }

    setQuestionDetails(): IQuestionState {
        let questionDetails: IQuestionState;
        if (this.props.questionMode == QuestionModeEnum.Create) {
            questionDetails = {
                questionMode: this.props.questionMode,
                questionText: "",
                questionType: QuestionTypesEnum.Text,
                answerText: "",
                required: true,
                choices: [],
                currentChoice: "",
                ratingLevel: 5,
                longTextAnswer: false,
                hideDeleteConfirmation: true
            }
        }
        else {            
            questionDetails = {
                questionMode: this.props.questionMode,
                questionText: this.props.questionText,
                questionType: this.props.questionType,
                answerText: this.props.answerText,
                required: this.props.required,
                choices: (this.props.choices != null && this.props.choices.length != 0) ? this.props.choices : [],
                currentChoice: "",
                ratingLevel: this.props.ratingLevel != null ? this.props.ratingLevel : 5,
                longTextAnswer: this.props.longTextAnswer,
                hideDeleteConfirmation: true
            }
        }
        return questionDetails;
    }

    deleteQuestionConfirmation(questionID: number): void {
        this.setState({
            hideDeleteConfirmation : false
        });
        //this.deleteQuestion(questionID);
    }

    closeDeleteQuestionConfirmation() : void{
        this.setState({
            hideDeleteConfirmation : true
        });
    }

    setEditQuestionMode(questionID: number): void {
        this.setState({
            questionMode: QuestionModeEnum.Edit
        });
    }

    saveQuestion(): void {
        if (this.state.questionMode == QuestionModeEnum.Create) {
            this.createQuestion();
        }
        else if (this.state.questionMode == QuestionModeEnum.Edit) {
            this.updateQuestion();
        }
    }

    createQuestion(): void {
        this.getListItemEntityTypeName()
            .then((listItemEntityTypeName: string): Promise<SPHttpClientResponse> => {
                const body: string = JSON.stringify({
                    '__metadata': {
                        'type': listItemEntityTypeName
                    },
                    'Title': this.state.questionText,
                    'Question_x0020_Type': this.state.questionType,
                    'Answer_x0020_Text': this.state.questionType == QuestionTypesEnum.Text ? this.state.answerText : "",
                    'Required': this.state.required,
                    'Choices': this.state.questionType == QuestionTypesEnum.Choice ? this.state.choices.join(";") : null,
                    'Rating_x0020_Level': this.state.questionType == QuestionTypesEnum.Rating ? this.state.ratingLevel : null,
                    'Is_x0020_Long_x0020_Text_x0020_A': this.state.questionType == QuestionTypesEnum.Text ? this.state.longTextAnswer : false
                });
                return this.props.spHttpClient.post(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items`,
                    SPHttpClient.configurations.v1,
                    {
                        headers: {
                            'Accept': 'application/json;odata=nometadata',
                            'Content-type': 'application/json;odata=verbose',
                            'odata-version': ''
                        },
                        body: body
                    });
            }).then(
                () => {
                    this.setState(this.setQuestionDetails());
                }
            )
            , (error: any): void => {
                console.log('Error while creating the item: ' + error);
            };
    }

    updateQuestion() {
        this.getListItemEntityTypeName().then((listItemEntityTypeName: string) => {
            this.getListItemEtagValue(this.props.questionID).then((etag: string) => {

                const body: string = JSON.stringify({
                    '__metadata': {
                        'type': listItemEntityTypeName
                    },
                    'Title': this.state.questionText,
                    'Question_x0020_Type': this.state.questionType,
                    'Answer_x0020_Text': this.state.questionType == QuestionTypesEnum.Text ? this.state.answerText : "",
                    'Required': this.state.required,
                    'Choices': this.state.questionType == QuestionTypesEnum.Choice ? this.state.choices.join(";") : null,
                    'Rating_x0020_Level': this.state.questionType == QuestionTypesEnum.Rating ? this.state.ratingLevel : null,
                    'Is_x0020_Long_x0020_Text_x0020_A': this.state.questionType == QuestionTypesEnum.Text ? this.state.longTextAnswer : false
                });

                this.props.spHttpClient.post(this.props.siteUrl + "/_api/web/lists/getbytitle('" + this.props.listName + "')/items(" + this.props.questionID + ")",
                    SPHttpClient.configurations.v1,
                    {
                        headers: {
                            'Accept': 'application/json;odata=nometadata',
                            'Content-type': 'application/json;odata=verbose',
                            'odata-version': '',
                            'IF-MATCH': etag,
                            'X-HTTP-Method': 'MERGE'
                        },
                        body: body
                    }).
                    then((response: SPHttpClientResponse): void => {
                        // On successful update, getting the all items
                        this.setState({questionMode : QuestionModeEnum.View});
                        this.props.onUpdate();
                    }, (error: any): void => {
                        // Failure
                    });
            });
        });
    }

    deleteQuestion(questionID: number) {
        this.getListItemEtagValue(questionID).then(
            (response: string) => {
                if (response != "") {
                    const etag = response;
                    this.props.spHttpClient.post(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${questionID})`,
                        SPHttpClient.configurations.v1,
                        {
                            headers: {
                                'Accept': 'application/json;odata=nometadata',
                                'Content-type': 'application/json;odata=verbose',
                                'odata-version': '',
                                'IF-MATCH': etag,
                                'X-HTTP-Method': 'DELETE'
                            }
                        }).then((response: SPHttpClientResponse): void => {
                            // On successful delete, getting the all items
                            this.props.onDelete();
                        }, (error: any): void => {
                            // On Failter
                        })
                }
                else {
                    alert("failed");
                }
            }
        );
    }

    getListItemEtagValue(itemID: number) {
        return new Promise<string>((resolve, reject) => {
            this.props.spHttpClient.get(this.props.siteUrl + "/_api/web/lists/getbytitle('" + this.props.listName + "')/items(" + itemID + ")?$select=Id",
                SPHttpClient.configurations.v1,
                {
                    headers: {
                        'Accept': 'application/json;odata=nometadata',
                        'odata-version': ''
                    }
                }).then(
                    // Success
                    (response: SPHttpClientResponse) => {
                        resolve(response.headers.get('ETag'));
                    },
                    // Failure
                    () => {
                        reject("");
                    });
        });
    }


    getListItemEntityTypeName(): Promise<string> {
        return new Promise<string>((resolve: (listItemEntityTypeName: string) => void, reject: (error: any) => void): void => {
            if (this.listItemEntityTypeName) {
                resolve(this.listItemEntityTypeName);
                return;
            }
            this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')?$select=ListItemEntityTypeFullName`,
                SPHttpClient.configurations.v1,
                {
                    headers: {
                        'Accept': 'application/json;odata=nometadata',
                        'odata-version': ''
                    }
                })
                .then(
                    // Success Function
                    (response: SPHttpClientResponse): Promise<{ ListItemEntityTypeFullName: string }> => {
                        return response.json();
                    },
                    // Error Function
                    (error: any): void => {
                        reject(error);
                    }
                )
                // Saving the result into class level variable
                .then((response: { ListItemEntityTypeFullName: string }): void => {
                    this.listItemEntityTypeName = response.ListItemEntityTypeFullName;
                    resolve(this.listItemEntityTypeName);
                });
        });
    }
}