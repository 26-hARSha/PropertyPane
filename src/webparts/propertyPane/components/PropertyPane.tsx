import * as React from 'react';
//import styles from './PropertyPane.module.scss';
import { IPropertyPaneProps } from './IPropertyPaneProps';
import styles from './PropertyPane.module.scss';
//import { escape } from '@microsoft/sp-lodash-subset';

export default class PropertyPane extends React.Component<IPropertyPaneProps, {}> {
  public render(): React.ReactElement<IPropertyPaneProps> {


    return (
      <div className={styles.welcome}>
       
        <h2 style={{color:'light blue'}}>Using Variables</h2>
        <h3>User Name:- {this.props.userDisplayName}</h3>
        <h3>Enviroment:- {this.props.environmentMessage}</h3>
        <h3>Site URL:- {this.props.siteAbsoluteURL}</h3>
        <h3>Site Title:- {this.props.siteTitle}</h3><br />

        <h2 style={{color:'blue'}}>User Properties(using Property pane)</h2>
        <h3>User Name:- {this.props.getUserName}</h3>
        <h3>User Age:- {this.props.getAge}</h3>
        <h3>User Hobbie:- {this.props.Hobbie}</h3>
        <h3>Is Married:- {" "} {this.props.IsMarried ?"Yes, Is Married":"No, Is Not" }</h3>
        <h3>Department:- {this.props.DropDown}</h3>
        <h3>Discount:- {" "} {this.props.Discount ? "YES":"No"}</h3>
      </div>
    );
  }
}
