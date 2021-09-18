import * as React from 'react';
import styles from './ZoneContentDisplay.module.scss';
import { IZoneContentDisplayProps } from './IZoneContentDisplayProps';
import { escape } from '@microsoft/sp-lodash-subset';

var oldSelectedList = "", oldSelectedZone = "";

export interface IZoneContentDisplayState {
  zoneContent: string;
}

export default class ZoneContentDisplay extends React.Component<IZoneContentDisplayProps, IZoneContentDisplayState> {

  constructor(props: IZoneContentDisplayProps, state: IZoneContentDisplayState) {
    super(props);

    this.state = {
      zoneContent: ""
    };
  }

  public componentDidMount() {
    this.getZoneContentToDisplay();
  }

  public componentDidUpdate() {
    if(this.props.selectedList != oldSelectedList || this.props.selectedZone != oldSelectedZone){
      this.getZoneContentToDisplay();
    }
  }

  public async getZoneContentToDisplay() {
    var apiUrl = this.props.siteAbsoluteURL + "/_api/web/lists/getbytitle('" + this.props.selectedList + "')/items?$select=*&$filter=Title eq '" + this.props.selectedZone + "'";

    oldSelectedList = this.props.selectedList;
    oldSelectedZone = this.props.selectedZone;

    await fetch(apiUrl, {
      headers: {
        credentials: "include",
        Accept: "application/json; odata=nometadata",
        "Content-Type": "application/json; odata=nometadata"
      }, method: "GET"
    })
      .then(response => response.json())
      .then((results) => {
        if (results.value != null && results.value.length > 0) {
          results.value.map((item) => {
            this.setState({ zoneContent: item.Content });
          });
        }
      });
  }


  public render(): React.ReactElement<IZoneContentDisplayProps> {
    return (
      <div className={styles.zoneContentDisplay}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <div>
                <p dangerouslySetInnerHTML={{ __html: this.state.zoneContent }}></p>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
