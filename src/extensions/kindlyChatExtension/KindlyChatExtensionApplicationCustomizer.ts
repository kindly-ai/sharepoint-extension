import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';


const LOG_SOURCE: string = 'KindlyChatExtensionApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IKindlyChatExtensionApplicationCustomizerProperties {
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class KindlyChatExtensionApplicationCustomizer
  extends BaseApplicationCustomizer<IKindlyChatExtensionApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initializing Kindly Chat`);
    const SCRIPT_ID = 'kindly-chat';
    const BOT_KEY = 'e08e78e7-d326-462e-9be9-cfd3b263248f';  // TODO: Update this

    if(document.getElementById(SCRIPT_ID)) {
      Log.info(LOG_SOURCE, `Kindly chat already loaded.`);
      return Promise.resolve();  // Kindly Chat is already initialized
    }
    const script = document.createElement("script");
    script.src = "https://chat.kindlycdn.com/kindly-chat-pure.js";
    script.async = true;
    script.id = 'kindly-chat';
    script.setAttribute('data-bot-key', BOT_KEY);
    script.setAttribute('data-shadow-dom', '1');
    script.setAttribute('data-bubble-avatar', 'https://i.imgur.com/3TpT5yV.png');
    document.head.appendChild(script);

    Log.info(LOG_SOURCE, `Appended script to head`);

    return Promise.resolve();
  }
}
