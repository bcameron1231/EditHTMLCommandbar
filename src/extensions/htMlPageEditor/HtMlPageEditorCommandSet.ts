import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'HtMlPageEditorCommandSetStrings';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IHtMlPageEditorCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
}

const LOG_SOURCE: string = 'HtMlPageEditorCommandSet';

export default class HtMlPageEditorCommandSet extends BaseListViewCommandSet<IHtMlPageEditorCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized HtMlPageEditorCommandSet');
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = false;
      if(event.selectedRows.length > 0)
      {
        compareOneCommand.visible = event.selectedRows.length === 1 && event.selectedRows[0].getValueByName(".fileType") == "html";
      }      
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':
      let fileName: string = event.selectedRows[0].getValueByName("FileLeafRef"); //Test.html,              
      let id: string = event.selectedRows[0].getValueByName("FileRef");
      let parent = id.substring(0,id.indexOf(fileName));
      location.search = "id=" + id + "&parent="+parent + "&p=5";
        break;   
      default:
        throw new Error('Unknown command');
    }
  }
}
