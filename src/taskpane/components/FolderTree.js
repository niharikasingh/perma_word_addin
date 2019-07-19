import * as React from 'react';
import { Icon } from 'office-ui-fabric-react/lib/Icon';

// Component that renders the user's folder structure

export default class FolderTree extends React.Component {
  constructor(props, context) {
    super(props, context);
    // details about current folder
    this.state = {
      "id": this.props.id,
      "name": this.props.name,
      "parent": this.props.parent,
      "has_children": this.props.has_children,
      "path": this.props.path,
      "organization": this.props.organization,
      "open": false,
      "children": null, // null means no request made yet to ask for all children of this folder
      "apiKey": this.props.apiKey
    }
  }

  toggleFolder = async () => {
    console.log("Clicked folder, started handler.");
    // we already have children folders, don't need to fetch them again
    if ((this.state.children != null) || (this.state.has_children == false)) {
      this.setState(prevState => ({
        "open": !prevState.open
      }));
    }
    else { // fetch children folders
      console.log("Clicked folder, fetching children.");
      fetch("/v1/folders/" + this.state.id + "/folders?api_key=" + this.state.apiKey, {
        method: "GET",
      }).then(async response => {
        if (response.status != 200) {
          throw "Error in toggleFolder response " + this.state.id;
        }
        // store children folders' details
        let responsebody = await response.json();
        console.log("fetch toggleFolder responsebody " + JSON.stringify(responsebody));
        this.setState(prevState => ({
          "children": responsebody.objects,
          "open": !prevState.open
        }));
      }).catch(err => {
        console.log(err);
      });
    }
  }

  render() {
    return (
      <div className="folder">
        {this.state.has_children &&
          <Icon
            className="folderIcon"
            iconName={this.state.open ? "Blocked2" : "CirclePlus"}
            onClick={this.toggleFolder}
            alt="Click to open or close this folder" />
        }
        {!this.state.has_children && <Icon className="folderIcon" iconName="LocationDot" />}
        <div
          className={"folderName" + ((this.props.selectedFolderId == this.state.id) ? " selected" : "")}
          onClick={() => this.props.updateFn(this.state.id)}>
          {this.state.name}
        </div>
        <div className="folderLevelDown">
          {this.state.open && !this.state.has_children &&
            <span>This folder is empty</span>
          }
          {this.state.open && this.state.has_children && // recursion!!!
            this.state.children.map(child_folder =>
              <FolderTree
                apiKey={this.state.apiKey}
                key={child_folder.id}
                id={child_folder.id}
                name={child_folder.name}
                parent={child_folder.parent}
                has_children={child_folder.has_children}
                path={child_folder.path}
                organization={child_folder.organization}
                updateFn={this.props.updateFn}
                selectedFolderId={this.props.selectedFolderId} />
            )
          }
        </div>
      </div>
    );
  }
}
