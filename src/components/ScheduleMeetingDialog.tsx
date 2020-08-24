import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import {
  PrimaryButton,
  TextField,
  DialogFooter,
  DialogContent
} from 'office-ui-fabric-react';
import styles from './ScheduleMeetingDialog.module.scss';
import { ExtensionContext } from '@microsoft/sp-extension-base';
import * as strings from 'DiscussNowCommandSetStrings';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IUserDetail } from '../models/OfficeUiFabricPeoplePicker';


const validated: boolean = true;

interface IScheduleMeetingDialogContentProps {
  context: any;
  siteURL: string;
  fileName: string;
  filePath: string;
  close: () => void;
  submit: (subject: string, siteURL: string, filePath: string, userEmails: string) => void;
}

interface IScheduleMeetingDialogContentState {
  subject: string;
  UserDetails: IUserDetail[];
  selectedusers: string[];
}

class ScheduleMeetingDialogContent extends React.Component<IScheduleMeetingDialogContentProps, IScheduleMeetingDialogContentState, {}> {

  constructor(props) {
    super(props);
    this.handleSubject = this.handleSubject.bind(this);
    this._getPeoplePickerItems = this._getPeoplePickerItems.bind(this);
    this._getErrorMessageSubject = this._getErrorMessageSubject.bind(this);
    this.state = {
      subject: this.props.fileName,
      UserDetails: [],
      selectedusers: []
    };
  }

  public render(): JSX.Element {
    return (<div className={styles.scheduleMeetingDialogRoot}>
      <DialogContent
        title={strings.DiscussNowDialogTitle}
        subText={strings.DiscussNowDialogDescription}
        onDismiss={this.props.close}
        showCloseButton={true}
      >
        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12">
            <PeoplePicker
              context={this.props.context}
              titleText="Select user(s)"
              personSelectionLimit={10}
              groupName={""} // Leave this blank in case you want to filter from all users    
              showtooltip={true}
              isRequired
              disabled={false}
              ensureUser={true}
              selectedItems={this._getPeoplePickerItems}
              defaultSelectedUsers={this.state.selectedusers}
              showHiddenInUI={false}
              principalTypes={[PrincipalType.User]}
              resolveDelay={1000}
            />

          </div>
        </div>
        <div className={styles.scheduleMeetingDialogContent}>
          <div className="ms-Grid">
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12">
                <TextField
                  label={strings.ScheduleMeetingSubjectLabel}                  
                  onChange={this.handleSubject}
                  value={this.state.subject}
                  multiline                
                />
              </div>
            </div>
          </div>
        </div>

        <DialogFooter>
          <PrimaryButton text='OK'
          disabled={this.state.selectedusers.length == 0}
          title='OK' onClick={() => { this.props.submit(this.state.subject, this.props.siteURL, this.props.filePath, this.state.selectedusers.join()); }} />
          <PrimaryButton text='Cancel' title='Cancel' onClick={this.props.close} />
        </DialogFooter>
      </DialogContent>
    </div>);
  }

  private _getErrorMessageSubject(value: string): string {
    return (value == null || value.length == 0 || value.length >= 10)
      ? ''
      : `${strings.ScheduleMeetingSubjectValidationErrorMessage} ${value.length}.`;
  }

  private _getPeoplePickerItems(items: any[]) {
    let userarr: IUserDetail[] = [];
    items.forEach(user => {
      userarr.push({ ID: user.id, LoginName: user.loginName });
    })

    let usernamearr: string[] = [];
    items.forEach(user => {
      usernamearr.push(user.loginName.split('|membership|')[1].toString());
    })
    this.setState({
      UserDetails: userarr,
      selectedusers: usernamearr
    })
  }

  private handleSubject(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void {
    this.setState({
      subject: newValue,
    });
  }
}

export default class ScheduleMeetingDialog extends BaseDialog {
  public fileName: string;
  public filePath: string;
  public context: ExtensionContext;

  public render(): void {
    ReactDOM.render(<ScheduleMeetingDialogContent
      context={this.context}
      siteURL={this.context.pageContext.site.absoluteUrl}
      fileName={this.fileName}
      filePath={this.filePath}
      close={this.close}
      submit={this._submit}
    />, this.domElement);
  }

  public getConfig(): IDialogConfiguration {
    return {
      isBlocking: false
    };
  }

  private async _submit(subject: string, siteURL: string, filePath: string, userEmails: string): Promise<void> {
    // *******************************************
    // Create deepkink for ms teams personal chat
    // *******************************************

    const url = "https://teams.microsoft.com/l/chat/0/0?users=" + userEmails + "&message=" + subject + "%0A" + encodeURI(filePath);
    window.open(url);
    this.close();
  }
}