import * as React from 'react';
import styles from './SpfxWebpart2.module.scss';
import type { ISpfxWebpart2Props } from './ISpfxWebpart2Props';
import { escape } from '@microsoft/sp-lodash-subset';

import { MSGraphClientV3 } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { DocumentCard, DocumentCardTitle, List, Spinner, SpinnerSize } from '@fluentui/react';
import * as strings from 'SpfxWebpart2WebPartStrings';


export default class SpfxWebpart2 extends React.Component<ISpfxWebpart2Props, { loading: boolean, messages: MicrosoftGraph.Message[], error : ''}> {

  constructor(props: ISpfxWebpart2Props){
    super(props);
    
    this.state = {
      loading: true,
      messages: [],
      error : ''
    }

    this._getMessages();
  }


  private _getMessages():void{
    this.props.context.msGraphClientFactory
    .getClient('3')
    .then((client:MSGraphClientV3):void =>{
      client
      .api('/me/messages')
      .top(5)
      .orderby('receivedDateTime desc')
      .get((error, messages, rawResponse?: string)=>{
        if(error) console.log(`[ERROR] ${error}`);
        else {
          this.setState({
            loading: false,
            messages : messages.value
          });
        }
      })
      .catch((e)=>{ console.error(`[OUPS 2] ${e}`);});
    })
    .catch((e)=>{ console.error(`[OUPS 1] ${e}`);});
  }


  public render(): React.ReactElement<ISpfxWebpart2Props> {
    const {
      description,
      //isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,      
    } = this.props;

    return (
      <section className={`${styles.spfxWebpart2} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          {/* <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} /> */}
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>

          {
          this.state.loading &&
            <Spinner label={strings.Loading} size={SpinnerSize.large} />
          }
          {
            this.state.messages && this.state.messages.length > 0 ? 
            (
              <List items={this.state.messages} onRenderCell={this._onRenderCell} />
            ):
            (
              !this.state.loading && (                
                <div>you have no messages</div>
              )
            )
          }

        </div>
        
      </section>
    );
  }

  private _onRenderCell = (item: MicrosoftGraph.Message, index: number | undefined): JSX.Element => {
    
    return <DocumentCard>
      <DocumentCardTitle title={item.subject ?? "Title is missing"} shouldTruncate={true} />
    </DocumentCard>

  }

}
