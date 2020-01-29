import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer, PlaceholderContent, PlaceholderName
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'GlobalMessageApplicationCustomizerStrings';
import styles from './GlobalMessageApplicationCustomizer.module.scss';

const LOG_SOURCE: string = 'GlobalMessageApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IGlobalMessageApplicationCustomizerProperties {
  // This is an example; replace with your own property
  TopMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class GlobalMessageApplicationCustomizer
  extends BaseApplicationCustomizer<IGlobalMessageApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    
    this.context.placeholderProvider.changedEvent.add(this, this.renderHeader);
	
    return Promise.resolve<void>();
  }

  private renderHeader(): void {
    
    let topPlaceholder: PlaceholderContent = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top); 

    if (topPlaceholder) {

      topPlaceholder.domElement.innerHTML = `<div class="${styles.app}">
      <div class="ms-bgColor-themeDark ms-fontColor-white ${styles.header}">
        <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i>&nbsp; ${this.properties.TopMessage}
      </div>
    </div>`;

    }
  }
}
