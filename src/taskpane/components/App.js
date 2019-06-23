import * as React from 'react';
import { PrimaryButton } from 'office-ui-fabric-react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { Separator } from 'office-ui-fabric-react/lib/Separator';
import Header from './Header';
import Progress from './Progress';

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
    this.state = {
      apiKey: "", 
      pending: []
    };
  }

  componentDidMount() {
    this.setState({
      apiKey: "", 
      pending: []
    })
  }

  showErrorMessage = (newValue) => {
    if (newValue.length == 40) {
      this.setState({apiKey: newValue});
      console.log("this.state");
      console.log(this.state);
      return "";
    }
    this.setState({apiKey: ""});
    console.log("this.state");
    console.log(this.state);
    return "API Key should be 40 characters"
  }

  click = async () => {
    return Word.run(async context => {
      // get current selection object
      console.log("starting click function");
      let currentCiteSelection = context.document.getSelection();
      // get current selection text 
      currentCiteSelection.load("text");
      await context.sync();
      let currentCiteSelectionText = currentCiteSelection.text;
      console.log("click: currentCiteSelectionText " + currentCiteSelectionText);
      // track perma requests that are pending 
      this.setState(prevState => ({pending: [...prevState.pending, currentCiteSelectionText]}));
      console.log("this.state");
      console.log(this.state);
      // save context of selection so it can be used in callback after fetch 
      context.trackedObjects.add(currentCiteSelection);
      // fetch perma.cc link for this selection
      fetch("https://cors-anywhere.herokuapp.com/https://api.perma.cc/v1/folders?api_key=" + this.state.apiKey, {
        method: "GET",
        mode: "cors",
        headers: {
          "Access-Control-Allow-Origin": "*"
        }
      }).then(async response => {
        let responsebody = await response.json();
        let folders = "";
        for (let f of responsebody.objects) {
          folders += f.name;
        }
        console.log("fetch folders " + folders);
        currentCiteSelection.insertText(folders, Word.InsertLocation.end);
        currentCiteSelection.context.trackedObjects.remove(currentCiteSelection);
        currentCiteSelection.context.sync();
        this.setState(prevState => (
          {pending: prevState.pending.filter(item => item !== currentCiteSelectionText)}
        ));
        console.log("this.state");
        console.log(this.state);
      }).catch(err => {
        console.log(err);
      });
      await context.sync();
    });
  }

  render() {
    const {
      title,
      isOfficeInitialized,
    } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo='assets/logo-filled.png'
          message='Please sideload your addin to see app body.'
        />
      );
    }

    return (
      <div className='ms-welcome'>
        <Header logo='assets/logo-filled.png' title={this.props.title} message='Welcome' />
        <div style={{padding: '20px'}}>
          <TextField label="Your API Key" onGetErrorMessage={this.showErrorMessage} validateOnLoad={false} />
          <p />
          <PrimaryButton iconProps={{ iconName: 'ChevronRight' }} onClick={this.click}>Run</PrimaryButton>
        </div>
        {this.state.pending.length > 0 &&
          <Separator>Currently pending</Separator>
        }
        {this.state.pending.map((p, i) => 
          <Spinner size={SpinnerSize.xSmall} labelPosition="right" label={p} key={i} />
        )}
      </div>
    );
  }
}
