import * as React from 'react';
import { Log, FormDisplayMode } from '@microsoft/sp-core-library';
import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';
import { WidgetSize, Dashboard } from '@pnp/spfx-controls-react/lib/Dashboard';
import "@pnp/sp/webs";
import "@pnp/sp/lists/web";
import "@pnp/sp/items";
import "@pnp/sp/attachments";
import { IAttachmentInfo, IEmailProperties, IItem, spfi, SPFx } from "@pnp/sp/presets/all";
import { DefaultButton, MessageBar, MessageBarType, ProgressIndicator } from '@fluentui/react';
import { getSP } from '../../../MyHelperMethods/MyHelperMethods';


export interface IByodProps {
  context: FormCustomizerContext;
  displayMode: FormDisplayMode;
  onSave: () => void;
  onClose: () => void;
}

const LOG_SOURCE: string = 'Byod';

export default class Byod extends React.Component<IByodProps, any> {
  /**
   *
   */
  constructor(props: IByodProps) {
    super(props);
    this._item = this.props.context.item;
    console.log(this.props);
    console.log('Item:');
    console.log(this.props.context.item);

    this.state = {
      listItemAttachments: undefined,
      isSidePanelOpen: false
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
    const item: IItem = sp.web.lists.getByTitle("BYOD Staff Agreement Submissions").items.getById(this._item.ID);

    // get all the attachments
    const info: IAttachmentInfo[] = await item.attachmentFiles();
    return info;
  }

  private _onCancelSubmissionClick = async (): Promise<any> => {


    let emailProps: IEmailProperties = {
      To: ['schorkawy@clarington.net'],
      Subject:'Test PnP Util Email',
      Body: 'This email was sent from an SPFx webpart without a workflow!  It will only work with internal Clarington.net accounts.'
    }
    await getSP().utility.sendEmail(emailProps);

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
                      onClick={this._onCancelSubmissionClick}
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
