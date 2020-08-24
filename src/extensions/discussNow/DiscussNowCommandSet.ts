import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import ScheduleMeetingDialog from '../../components/ScheduleMeetingDialog';
import * as strings from 'DiscussNowCommandSetStrings';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IDiscussNowCommandSetProperties {
}

const LOG_SOURCE: string = 'DiscussNowCommandSet';

export default class DiscussNowCommandSet extends BaseListViewCommandSet<IDiscussNowCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized DiscussNowCommandSet');
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('DISCUSS_NOW');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = event.selectedRows.length === 1;
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'DISCUSS_NOW':
        // const id: number = event.selectedRows[0].getValueByName("ID");
        const fileType = event.selectedRows[0].getValueByName("File_x0020_Type") == "" ? "List" : "Libraray";
        const fileName: string = fileType == "List" ? event.selectedRows[0].getValueByName("Title") : event.selectedRows[0].getValueByName("FileLeafRef");
        const filePath: string = fileType == "List" ? this.context.pageContext.site.absoluteUrl.substring(0, this.context.pageContext.site.absoluteUrl.indexOf('/sites')) + this.context.pageContext.list.serverRelativeUrl + "/DispForm.aspx?ID=" + event.selectedRows[0].getValueByName("ID")
          : this.context.pageContext.site.absoluteUrl.substring(0, this.context.pageContext.site.absoluteUrl.indexOf('/sites')) + event.selectedRows[0].getValueByName("FileRef") + "?web=1";
        
          const dialog: ScheduleMeetingDialog = new ScheduleMeetingDialog();
        dialog.fileName = fileName;
        dialog.filePath = filePath;
        dialog.context = this.context;
        dialog.show();
        break;

      default:
        throw new Error('Unknown command');
    }
  }
}
