import * as React from 'react'
import { ITextProps } from './ITextProps';

export class TextComponent extends React.Component<ITextProps, {}>{
    render(): React.ReactElement<ITextProps> {
        return (
            <div>
                <input
                    type="text"
                    value={this.props.answerText}
                    placeholder={this.props.placeholderContent}
                    maxLength={this.props.maxChars}
                    minLength={this.props.minChars}>
                </input>
            </div>
        );
    }
}
