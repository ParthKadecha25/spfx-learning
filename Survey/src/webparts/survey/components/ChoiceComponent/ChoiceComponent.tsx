import * as React from 'react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { ActionButton, IconButton } from 'office-ui-fabric-react/lib/Button';
import { IChoiceProps } from './IChoiceProps'
import styles from './Choice.module.scss'
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { IChoiceState } from './IChoiceState';

export class ChoiceComponent extends React.Component<IChoiceProps,IChoiceState>{
    
    constructor(props){
        super(props);
        this.state = {
            choice : ""
        }
    }

    render() : JSX.Element {
        console.log(this.props.choices);
        const prevChoiceEntries = this.props.choices.map( (choice) => 
                                    <div className="ms-Grid-row">
                                        <div className="ms-Grid-col ms-sm1">
                                            <Icon iconName="RadioBtnOff" className={styles.radioIconStyle}/>
                                        </div>
                                        <div className="ms-Grid-col ms-sm9">
                                            <TextField type="text" value={choice} readOnly={true}></TextField>
                                        </div>
                                        <div className="ms-Grid-col ms-sm2">
                                            <IconButton iconProps={{ iconName: 'delete' }} title="Remove" ariaLabel="Remove"/>
                                        </div>
                                    </div>
                                );
        const placeholder = "Choice " + (this.props.choices.length + 1)
        return(
            <div className="ms-Grid">
                {prevChoiceEntries}
                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm1">
                        <Icon iconName="RadioBtnOff" className={styles.radioIconStyle}/>
                    </div>
                    <div className="ms-Grid-col ms-sm9">
                        <TextField type="text" placeholder={placeholder} onChanged={ (value) => {this.setState({choice : value})}} ></TextField>
                    </div>
                </div>     
                <div className="ms-Grid-row">
                    <ActionButton text="Add Choice" iconProps={{iconName : "add"}} 
                        className={styles.addChoiceButton}
                        onClick={this.props.onAddChoice.bind(this, this.state.choice)}
                    />
                </div>
            </div>
        )
    }
}