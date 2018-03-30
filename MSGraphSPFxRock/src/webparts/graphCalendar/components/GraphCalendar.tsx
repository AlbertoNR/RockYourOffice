import * as React from 'react';
import styles from './GraphCalendar.module.scss';
import { IGraphCalendarProps } from './IGraphCalendarProps';
import { IGraphCalendarState } from './IGraphCalendarState';
import { escape } from '@microsoft/sp-lodash-subset';
import { MSGraphClient } from '@microsoft/sp-client-preview';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';


export default class GraphCalendar extends React.Component<IGraphCalendarProps, IGraphCalendarState> {

  constructor(props)
  {
    super(props);

    this.state = {
      user: '',
      event: {
        subject:'',
        bodyPreview:'',
        location: {displayName:''},
        start:{dateTime:''},
        end:{dateTime:''}
      }
    };

  }

  public componentDidMount(){
    const client: MSGraphClient = this.props.spContext.serviceScope.consume(MSGraphClient.serviceKey);
    // get information about the current user from the Microsoft Graph
    client
      .api('/me')
      .get((error, user: MicrosoftGraph.User, rawResponse?: any) => {

        console.log('Error: ' + error);

        this.setState({ user: user.displayName});
        console.log(user);
    });


    client
      .api('/me/events?$select=subject,body,bodyPreview,organizer,attendees,start,end,location')
      .get((error, response: any, rawResponse?: any) => {

        let events:[MicrosoftGraph.Event] = response.value;

        console.log(events);
        if(events.length>0){
           this.setState({ event: events[0] });
        }
    });
  }


  public render(): React.ReactElement<IGraphCalendarProps> {
    return (
      <div className={ styles.graphCalendar }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Bienvenido a SharePoint Folk!</span>
              <p className={ styles.subTitle }>{'Prueba de concepto con Microsoft Graph & SPFx WebPart.'}</p>
              <p className={ styles.description }>{ 'Usuario (API de MS Graph): ' + this.state.user }</p>
              <h2 className={ styles.description }>{ 'Evento (API de MS Graph):' }</h2>
              <ul>
                <li>{ `Subject: ${this.state.event.subject}` }</li>
                <li>{ `BodyPreview: ${this.state.event.bodyPreview}` }</li>
                <li>{ `Location: ${this.state.event.location.displayName}` }</li>
                <li>{ `Start: ${this.state.event.start.dateTime}` }</li>
                <li>{ `End: ${this.state.event.end.dateTime}` }</li>
              </ul>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
