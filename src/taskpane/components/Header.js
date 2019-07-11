import * as React from 'react';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';

export default class Header extends React.Component {
  render() {
    const {
      title,
      logo,
      message
    } = this.props;

    return (
      <section className='ms-welcome__header ms-bgColor-neutralLighter ms-u-fadeIn500'>
        <h3 className='ms-fontSize-su ms-fontWeight-light ms-fontColor-neutralPrimary'>{message}</h3>
          {this.props.showError && 
            <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
              {this.props.showErrorMessage}
            </MessageBar>
          }
      </section>
    );
  }
}
