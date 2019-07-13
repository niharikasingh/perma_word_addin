import * as React from 'react';
import { PrimaryButton } from 'office-ui-fabric-react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { Separator } from 'office-ui-fabric-react/lib/Separator';
import Header from './Header';
import Progress from './Progress';
import FolderTree from './FolderTree';

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
    this.state = {
      apiKey: "", // stores user's api key 
      apiKeyError: false, // True when api key does not lead to valid user details
      pending: [], // stores links that are pending transformation to perma links 
      userName: '', // stores user's full name 
      folderTree: [], // stores user's top-level folder tree
      selectedFolderId: '', // stores the folder ID to save links to 
    };
  }

  // update what the currently selected folder is
  updateSelectedFolderId = (id) => {
    console.log("updateSelectedFolderId function called with id: " + id);
    if (this.state.selectedFolderId != id) {
      this.setState({
        selectedFolderId: id
      })
    }
  }

  // validate API key by checking its length
  showErrorMessage = (newValue) => {
    if (newValue.length == 40) {
      this.setState({ apiKey: newValue });
      console.log("this.state");
      console.log(this.state);
      this.userDetails();
      return "";
    }
    this.setState({ apiKey: '', apiKeyError: false, userName: '' });
    console.log("this.state");
    console.log(this.state);
    return "API Key should be 40 characters"
  }

  // is API key valid?
  apiKeyNotReady = () => !(this.state.apiKey.length == 40) || this.state.apiKeyError;

  userDetails = async () => {
    // fetch user full name 
    fetch("https://cors-anywhere.herokuapp.com/https://api.perma.cc/v1/user/?api_key=" + this.state.apiKey, {
      method: "GET",
      mode: "cors",
      headers: {
        "Access-Control-Allow-Origin": "*",
        "Content-Type": "application/json",
        "Accept": "application/json",
      }
    }).then(async response => {
      if (response.status != 200) {
        throw "Error in userDetails response";
      }
      // get user details  
      let responsebody = await response.json();
      console.log("fetch userDetails responsebody " + JSON.stringify(responsebody));
      this.setState({
        apiKeyError: false,
        userName: responsebody.full_name,
        folderTree: responsebody.top_level_folders,
        selectedFolderId: responsebody.top_level_folders[0].id, // initialize to default folder = first top folder
      });
    }).catch(err => {
      console.log(err);
      // show apiKeyError
      this.setState({ apiKeyError: true, userName: '' });
    });
  }

  // upon click, return perma link of selected link text to user 
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
      this.setState(prevState => ({ pending: [...prevState.pending, currentCiteSelectionText] }));
      console.log("this.state");
      console.log(this.state);
      // save context of selection so it can be used in callback after fetch 
      context.trackedObjects.add(currentCiteSelection);
      // fetch perma.cc link for this selection
      fetch("https://cors-anywhere.herokuapp.com/https://api.perma.cc/v1/archives?api_key=" + this.state.apiKey, {
        method: "POST",
        mode: "cors",
        headers: {
          "Access-Control-Allow-Origin": "*",
          "Content-Type": "application/json",
          "Accept": "application/json",
        },
        body: JSON.stringify({ url: currentCiteSelectionText, folder: this.state.selectedFolderId })
      }).then(async response => {
        // get guid from response and insert into word document 
        let responsebody = await response.json();
        console.log("fetch guid " + responsebody.guid);
        currentCiteSelection.insertText(" [https://perma.cc/" + responsebody.guid + "]", Word.InsertLocation.end);
        // remove context of selection because we're done using it 
        currentCiteSelection.context.trackedObjects.remove(currentCiteSelection);
        currentCiteSelection.context.sync();
        // remove request from list of pending requests 
        this.setState(prevState => (
          { pending: prevState.pending.filter(item => item !== currentCiteSelectionText) }
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
        <Header
          logo='assets/logo-filled.png' title={this.props.title} message={this.state.userName} showError={this.state.apiKeyError} showErrorMessage="Your API Key does not seem to be valid." />

        <div className="folderTop">
          {!this.apiKeyNotReady() && this.state.folderTree.map(top_folder => {
            console.log("Top level folder being prepared for display: " + JSON.stringify(top_folder));
            return <FolderTree
              apiKey={this.state.apiKey}
              key={top_folder.id}
              id={top_folder.id}
              name={top_folder.name}
              parent={top_folder.parent}
              has_children={top_folder.has_children}
              path={top_folder.path}
              organization={top_folder.organization}
              updateFn={this.updateSelectedFolderId}
              selectedFolderId={this.state.selectedFolderId} />
          })}
        </div>

        <div style={{ padding: '20px' }}>
          <TextField label="Your API Key" onGetErrorMessage={this.showErrorMessage} validateOnLoad={false} />
          <p />
          <PrimaryButton
            disabled={this.apiKeyNotReady()}
            iconProps={{ iconName: 'ChevronRight' }} onClick={this.click}
          >
            Run
          </PrimaryButton>
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
