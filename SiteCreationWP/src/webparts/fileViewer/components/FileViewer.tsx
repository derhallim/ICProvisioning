import * as React from 'react';
import styles from './FileViewer.module.scss';
import { IFileViewerProps } from './IFileViewerProps';

var oldSelectedLibrary: string = "";

export interface IFileViewerState {
  selectedFileLibraryUrl: string;
}
export default class FileViewer extends React.Component<IFileViewerProps, IFileViewerState> {

  constructor(props: IFileViewerProps, state: IFileViewerState) {
    super(props);
    this.state = {
      selectedFileLibraryUrl: ""
    };
  }

  public componentDidMount() {

    setTimeout(() => {
      this.GetLibraryFileUrl();
    }, 3000);

  }

  public componentDidUpdate() {
    if (this.props.selectedLibrary != oldSelectedLibrary) {
      this.GetLibraryFileUrl();
    }
  }

  private GetLibraryFileUrl(): any {
    oldSelectedLibrary = this.props.selectedLibrary;
    var url = `${this.props.siteAbsoluteURL}/_api/web/lists/getbytitle('${this.props.selectedLibrary}')/items?$select=*&$filter=Status eq 'Active'&$orderby=Modified desc&$top=1`;

    fetch(url, {
      credentials: 'same-origin',
      headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' }
    })
      .then((res) => res.json())
      .then(
        (result) => {
          if (result.value != undefined) {
            this.setState({ selectedFileLibraryUrl: result.value[0].ServerRedirectedEmbedUri });
          }
        },
        (error) => {
          console.log(error);
        }
      );
  }

  public render(): React.ReactElement<IFileViewerProps> {
    return (
      <div className={styles.fileViewer}>
        <div>
          <iframe
            height={500}
            width="100%"
            src={this.state.selectedFileLibraryUrl}
          ></iframe>
        </div>
      </div>
    );
  }
}
