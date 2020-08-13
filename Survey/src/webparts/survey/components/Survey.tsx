import * as React from 'react';
import { ISurveyProps } from './ISurveyProps';
import { QuestionComponent } from './QuestionComponent/QuestionComponent';
import { ISurveyState } from './ISurveyState';
import styles from './Survey.module.scss';
import { Pivot, PivotItem, PivotLinkFormat, PivotLinkSize } from 'office-ui-fabric-react/lib/Pivot';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { QuestionModeEnum } from './QuestionModeEnum';
import { PrimaryButton, IButtonProps, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';

export default class Survey extends React.Component<ISurveyProps, ISurveyState> {

  private listItemEntityTypeName: string = undefined;

  constructor(props) {
    super(props);
    this.state = {
      questions: [],
      questionsCount: 0,
      questionsCountLoaded : false,
      fillSurveyBtnClicked : false,
    }
    this.getListItems = this.getListItems.bind(this);
    this.handleMenuItemClick = this.handleMenuItemClick.bind(this);
    this.handleFillSurveyBtnClick = this.handleFillSurveyBtnClick.bind(this);
    this.cancelSurveyResponse = this.cancelSurveyResponse.bind(this);
  }

  componentDidMount(){
    if(!this.props.adminMode){
      this.getListItems(this.props.listNameSurvey);
    }
  }

  render(): JSX.Element {
    return (
      <div className={styles.survey}>
        { this.props.adminMode ? (
          // CREATE SURVEY MODE
        <Pivot linkFormat={PivotLinkFormat.tabs} linkSize={PivotLinkSize.large} onLinkClick={this.handleMenuItemClick}>
          <PivotItem linkText="Add Question" itemIcon="add">
            <QuestionComponent
              questionMode={QuestionModeEnum.Create}
              listName={this.props.listNameSurvey}
              siteUrl={this.props.siteUrl}
              spHttpClient={this.props.spHttpClient}
              adminMode={this.props.adminMode}
            >
            </QuestionComponent>
          </PivotItem>
          <PivotItem linkText="View Questions" itemIcon="view" >
            <br />
            {
              this.state.questions.map((q, index) =>
                <QuestionComponent
                  key = {q.questionID}
                  questionMode={QuestionModeEnum.View}
                  listName={this.props.listNameSurvey}
                  siteUrl={this.props.siteUrl}
                  spHttpClient={this.props.spHttpClient}
                  answerText={q.answerText}
                  choices={q.choices}
                  longTextAnswer={q.longTextAnswer}
                  questionText={q.questionText}
                  questionType={q.questionType}
                  ratingLevel={q.ratingLevel}
                  required={q.required}
                  questionID={q.questionID}
                  questionNo={(index + 1)}
                  onDelete = {this.getListItems}
                  onUpdate = {this.getListItems}
                  adminMode = {this.props.adminMode}
                >
                </QuestionComponent>
              )
            }
            {this.state.questions.length == 0 && <h1>No Questions added yet!</h1>}
          </PivotItem>
        </Pivot>
        ) : (
          // FILL SURVEY MODE
          <div>
            <h1>
                Total Questions :
                {
                  this.state.questionsCountLoaded ? (
                    " " + this.state.questionsCount
                  ):
                  (
                    <Spinner size={SpinnerSize.large}/>
                  )
                }
            </h1>
            {!this.state.fillSurveyBtnClicked  && 
              (   
                <PrimaryButton text="Fill Survey" onClick={this.handleFillSurveyBtnClick}></PrimaryButton>
              )
            }
            {this.state.fillSurveyBtnClicked && 
              (   
                  this.state.questions.map((q, index) =>
                    <QuestionComponent
                      key = {q.questionID}
                      questionMode={QuestionModeEnum.View}
                      listName={this.props.listNameSurvey}
                      siteUrl={this.props.siteUrl}
                      spHttpClient={this.props.spHttpClient}
                      answerText={q.answerText}
                      choices={q.choices}
                      longTextAnswer={q.longTextAnswer}
                      questionText={q.questionText}
                      questionType={q.questionType}
                      ratingLevel={q.ratingLevel}
                      required={q.required}
                      questionID={q.questionID}
                      questionNo={(index + 1)}
                      onDelete = {this.getListItems}
                      onUpdate = {this.getListItems}
                      adminMode = {this.props.adminMode}
                    >
                    </QuestionComponent>
                  )
              )
            }
            {this.state.fillSurveyBtnClicked  && 
              (    
                <div>
                  <br/>
                  <PrimaryButton text="Save" onClick={this.handleFillSurveyBtnClick}></PrimaryButton>
                  <DefaultButton text="Cancel" onClick={this.cancelSurveyResponse}></DefaultButton>
                </div>
              )
            }
          </div>
        )}
      </div>
    );
  }

  handleFillSurveyBtnClick(): void
  {
    this.setState({
      fillSurveyBtnClicked : true
    });
    this.getListItems(this.props.listNameSurvey);
  }

  handleMenuItemClick(item: PivotItem): void {
    if (item.props.linkText == "View Questions") {
      this.getListItems(this.props.listNameSurvey);
    }
  }

  saveSurveyResponse() : void
  {
    // this.getListItemEntityTypeName()
    // .then((listItemEntityTypeName: string): Promise<SPHttpClientResponse> => {
    //     const body: string = JSON.stringify({
    //         '__metadata': {
    //             'type': listItemEntityTypeName
    //         },
    //         'Title': this.state.questionText,
    //         'Question_x0020_Type': this.state.questionType,
    //         'Answer_x0020_Text': this.state.questionType == QuestionTypesEnum.Text ? this.state.answerText : ""
    //     });
    //     return this.props.spHttpClient.post(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items`,
    //         SPHttpClient.configurations.v1,
    //         {
    //             headers: {
    //                 'Accept': 'application/json;odata=nometadata',
    //                 'Content-type': 'application/json;odata=verbose',
    //                 'odata-version': ''
    //             },
    //             body: body
    //         });
    // }).then(
    //     () => {
    //         this.setState(this.setQuestionDetails());
    //     }
    // )
    // , (error: any): void => {
    //     console.log('Error while creating the item: ' + error);
    // };
  }

  getListItemEntityTypeName(listName : string): Promise<string> {
    return new Promise<string>((resolve: (listItemEntityTypeName: string) => void, reject: (error: any) => void): void => {
        if (this.listItemEntityTypeName) {
            resolve(this.listItemEntityTypeName);
            return;
        }
        this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`,
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
  cancelSurveyResponse() : void
  {
      this.setState({
        fillSurveyBtnClicked : false
      })
  }

  getListItems(listName : string): Promise<any> {
    return this.props.spHttpClient.get(this.props.siteUrl + `/_api/web/lists/GetByTitle('` + listName + `')/Items`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      }).then((response) => {
        let items = response.value;
        let allQuestions = [];
        items.forEach((item) => {
          let question = {
            questionText: item["Title"],
            questionType: item["Question_x0020_Type"],
            answerText: item["Answer_x0020_Text"] == null ? "" : item["Answer_x0020_Text"],
            required: item["Required"],
            choices: item["Choices"] == null ? "" : item["Choices"].toString().split(";"),
            ratingLevel: item["Rating_x0020_Level"],
            longTextAnswer: item["Is_x0020_Long_x0020_Text_x0020_A"],
            questionID: item["ID"]
          }
          allQuestions.push(question);
        });
        this.setState({ 
          questions: allQuestions
         });
      });
  }

  getListItemCount(listName : string): Promise<any> {
    return this.props.spHttpClient.get(this.props.siteUrl + `/_api/web/lists/GetByTitle('` + listName + `')//ItemCount`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      }).then((response) => {
        if(response != null){
          this.setState({
            questionsCount: response.value,
            questionsCountLoaded : true
           });
        }
      });
  }
}
