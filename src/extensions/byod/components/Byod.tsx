import * as React from 'react';
import { Log, FormDisplayMode } from '@microsoft/sp-core-library';
import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';
import { WidgetSize, Dashboard } from '@pnp/spfx-controls-react/lib/Dashboard';
import "@pnp/sp/webs";
import "@pnp/sp/lists/web";
import "@pnp/sp/items";
import "@pnp/sp/attachments";
import { IAttachmentInfo, IEmailProperties, IItem, spfi, SPFx } from "@pnp/sp/presets/all";
import { DefaultButton, Link, MessageBar, MessageBarType, Panel, PrimaryButton, ProgressIndicator, TextField } from '@fluentui/react';
import { getSP } from '../../../MyHelperMethods/MyHelperMethods';


export interface IByodProps {
  context: FormCustomizerContext;
  displayMode: FormDisplayMode;
  onSave: () => void;
  onClose: () => void;
}

export interface IMyProgressIndicator {
  description: string;
  percentComplete: number;
}
export interface IByodState {
  isSidePanelOpen: boolean;
  listItemAttachments?: any;
  cancelationComments?: string;
  myProgressIndicator?: IMyProgressIndicator;
  showCancellationSuccessMessage: boolean;
}

const LOG_SOURCE: string = 'Byod';
const BYOD_LIST_TITLE = 'BYOD Staff Agreement Submissions';

export default class Byod extends React.Component<IByodProps, IByodState> {
  /**
   *
   */
  constructor(props: IByodProps) {
    super(props);
    this._item = this.props.context.item;

    this.state = {
      listItemAttachments: undefined,
      isSidePanelOpen: false,
      showCancellationSuccessMessage: false
    };

    if (this._item.Attachments) {
      this._getAttachments().then(value => {
        this.setState({
          listItemAttachments: value
        });
      });
    }
  }

  private _getAttachments = async (): Promise<any> => {
    const sp = spfi().using(SPFx(this.props.context as any));
    const item: IItem = sp.web.lists.getByTitle(BYOD_LIST_TITLE).items.getById(this._item.ID);

    // get all the attachments
    const info: IAttachmentInfo[] = await item.attachmentFiles();
    return info;
  }

  private _cancelRequest = async (): Promise<void> => {
    this.setState({ myProgressIndicator: { description: 'Getting Current User.', percentComplete: 10 } });
    const CURRENT_USER = await getSP().web.currentUser();

    this.setState({ myProgressIndicator: { description: 'Updating Metadata.', percentComplete: 50 } });
    await getSP().web.lists.getByTitle(BYOD_LIST_TITLE).items.getById(this._item.ID)
      .update({
        "OData__Status": 'Cancelled',
        "ApprovalComments": `${this._item.ApprovalComments}${CURRENT_USER.Title} - Cancelled - ${this.state.cancelationComments}\n\n`,
        "ApprovalSummary": `${this._item.ApprovalSummary}\nApprover: ${CURRENT_USER.Title}, ${CURRENT_USER.Email}\nResponse: Cancelled\nCancelled Date: ${new Date().toLocaleString()}\n`
      });
  }

  private _sendCancelationEmail = async (): Promise<void> => {
    const TO_EMAIL_PROPS = ['payroll@clarington.net'];
    this.setState({ myProgressIndicator: { description: `Sending Email to ${TO_EMAIL_PROPS}.`, percentComplete: 75 } });
    const CURRENT_USER = await getSP().web.currentUser();
    const BODY_EMAIL_PROPS = `
    <div>
    <h2>Notice of BYOD Cancelation</h2>
    <div>Please cancel the BYOD plan for "${this._item.Title}".</div>
    <a href="https://claringtonnet.sharepoint.com/sites/InfoTech/_layouts/15/SPListForm.aspx?PageType=4&List=e51fc631%2Dc4da%2D4847%2D8e50%2Da933571c0811&ID=${this._item.ID}">Click Here to View BYOD Submission Details.</a>
    <br/><br/>
    <div>Cancelation Comments from ${CURRENT_USER.Title}:</div>
    <div>"${this.state.cancelationComments}"</div>
    </div>
    `;
    const SUBJECT_EMAIL_PROPS = `BYOD Cancelation - ${this._item.Title}`

    let emailProps: IEmailProperties = {
      To: TO_EMAIL_PROPS,
      Subject: SUBJECT_EMAIL_PROPS,
      Body: BODY_EMAIL_PROPS
    }

    await getSP().utility.sendEmail(emailProps).catch(reason => { alert('Failed to send') });
    this.setState({ myProgressIndicator: { description: `Done!`, percentComplete: 100 }, showCancellationSuccessMessage: true });
  }

  private _updateRecordAndSendCancelationEmail = async (): Promise<void> => {
    await this._cancelRequest();
    await this._sendCancelationEmail();
  }

  private _item: any;

  public componentDidMount(): void {
    Log.info(LOG_SOURCE, 'React Element: Byod mounted');
  }

  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: Byod unmounted');
  }

  public render(): React.ReactElement<{}> {
    return <div>
      <Panel
        headerText="BYOD Cancelation Form"
        isOpen={this.state.isSidePanelOpen}
        closeButtonAriaLabel="Close"
        onDismiss={() => this.setState({ isSidePanelOpen: false })}
        onRenderFooterContent={() => {
          return <div>
            <p>Saving this form will automatically send an email to Payroll@clarington.net notifying them of the change.</p>
            <PrimaryButton onClick={this._updateRecordAndSendCancelationEmail} styles={{ root: { marginRight: 8 } }} disabled={this.state.myProgressIndicator !== undefined}>Save</PrimaryButton>
            <DefaultButton onClick={() => this.setState({ isSidePanelOpen: false })} disabled={this.state.myProgressIndicator !== undefined}>Close</DefaultButton>
            {
              (this.state.myProgressIndicator && this.state.showCancellationSuccessMessage === false) &&
              <div>
                <ProgressIndicator label="Saving Cancelation..." description={this.state.myProgressIndicator.description} percentComplete={this.state.myProgressIndicator.percentComplete} />
              </div>
            }
            {
              this.state.showCancellationSuccessMessage &&
              <MessageBar messageBarType={MessageBarType.success} isMultiline={true} style={{ marginTop: 5 }}>
                Success! The Status has been updated and Payroll has been notified.
                <Link href={window.location.href} underline>
                  Click Here to View Your Changes.
                </Link>
              </MessageBar>
            }

          </div>;
        }}

      >
        <TextField
          label="Cancelation Comments"
          multiline
          rows={6}
          onChange={(event, newValue) => this.setState({ cancelationComments: newValue })}
        />
      </Panel>
      <Dashboard
        widgets={
          [{
            title: "BYOD Submission",
            // widgetActionGroup: calloutItemsExample,
            size: WidgetSize.Box,
            body: [
              {
                id: "t1",
                title: "Tab 1",
                content: (
                  <div>
                    <h1>{this._item.Title}</h1>
                    <p>Submission Date: {this._item.Date1 && new Date(this._item.Date1).toLocaleDateString()}</p>
                    <p>Final Approval Date: {this._item.FinalApprovalDate && new Date(this._item.FinalApprovalDate).toLocaleDateString()}</p>
                    <p>Device/ Model: {this._item.DeviceManufacturer}/ {this._item.DeviceModel}</p>
                    <p>Operating System: {this._item.DeviceOperatingSystem}</p>
                    <p>Wireless Provider: {this._item.WirelessProvider}</p>
                    <p>Contract Type: {this._item.ContractType}</p>
                    <p>Contract End Date: {this._item.ContractEndDate && new Date(this._item.ContractEndDate).toLocaleDateString('en-US')}</p>
                    {
                      (this.state.listItemAttachments === undefined && this._item.Attachments === true) &&
                      <div>
                        <ProgressIndicator label="Loading Attachments..." />
                      </div>
                    }
                    {
                      (this.state.listItemAttachments !== undefined && this._item.Attachments === true) &&
                      <div>
                        <p>Attachments:</p>
                        {
                          this.state.listItemAttachments.map((attachment: any) => {
                            return <div><a href={attachment.ServerRelativeUrl}>{attachment.FileName}</a></div>;
                          })
                        }
                      </div>
                    }
                    {
                      (this.state.listItemAttachments === undefined && this._item.Attachments === false) && <div>
                        <MessageBar
                          messageBarType={MessageBarType.error}
                          isMultiline={false}
                        >
                          No Attachments Found...
                        </MessageBar>
                      </div>
                    }
                  </div>
                ),
              },
            ]
          },
          {
            title: "Approval Status",
            size: WidgetSize.Box,
            body: [
              {
                id: 'c2t1',
                title: 'Card 2 Title',
                content: (
                  <div>
                    <h1>{this._item.OData__Status}</h1>
                    <DefaultButton
                      text="Cancel Submission"
                      iconProps={{ iconName: 'Blocked' }}
                      onClick={() => this.setState({ isSidePanelOpen: true })}
                    />
                    <hr />
                    <h4>Approval Comments</h4>
                    <pre>{this._item.ApprovalComments}</pre>
                    <hr />
                    <h4>Approval Summary</h4>
                    <pre>{this._item.ApprovalSummary}</pre>
                  </div>
                )
              }
            ]
          }
          ]} />
    </div >;
  }
}
