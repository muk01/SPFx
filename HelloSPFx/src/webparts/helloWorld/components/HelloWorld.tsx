import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
import { escape, fromPairs } from '@microsoft/sp-lodash-subset';

import {IHelloWorldState} from './IHelloWorldState';
import {IListItem} from './IListItem';
import { SPHttpClient, SPHttpClientConfiguration, SPHttpClientResponse, 
  ISPHttpClientConfiguration } from '@microsoft/sp-http';

export default class HelloWorld extends React.Component<IHelloWorldProps, IHelloWorldState> {

  constructor (props: IHelloWorldProps, state:IHelloWorldState){
    super(props);
    this.state={
      status:'Ready',
      Items:[]
    };
  }

  public render(): React.ReactElement<IHelloWorldProps> {

    const items: JSX.Element[] = this.state.Items.map((item: IListItem, i:number): JSX.Element => {
      return (
      <li>{item.Title} {item.Id}</li>
      );
    });

    return (
      <div className={ styles.helloWorld }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.ListName)}</p>
              
              <div className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`}>  
                <div className='ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1'>  
                  <a href="#" className={`${styles.button}`} onClick={() => this.createItem()}>  
                    <span className={styles.label}>Create item</span>  
                  </a>   
                  <a href="#" className={`${styles.button}`} onClick={() => this.readItem()}>  
                    <span className={styles.label}>Read item</span>  
                  </a>  
                </div>  
              </div>  
  
              <div className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`}>  
                <div className='ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1'>  
                  <a href="#" className={`${styles.button}`} onClick={() => this.updateItem()}>  
                    <span className={styles.label}>Update item</span>  
                  </a>   
                  <a href="#" className={`${styles.button}`} onClick={() => this.deleteItem()}>  
                    <span className={styles.label}>Delete item</span>  
                  </a>  
                </div>  
              </div>  
  
              <div className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`}>  
                <div className='ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1'>  
                  {this.state.status}  
                  <ul>  
                    {items}  
                  </ul>  
                </div>  
              </div> 

            </div>
          </div>
        </div>
      </div>
    );
  }

  private createItem(): void {  
    this.setState({  
      status: 'Creating item...',  
      Items: []  
    });  
    
    const body: string = JSON.stringify({  
      'Title': `Item ${new Date()}`  
    });  
    
    this.props.spHttpClient.post(`${this.props.siteURL}/_api/web/lists/getbytitle('${this.props.ListName}')/items`,  
    SPHttpClient.configurations.v1,  
    {  
      headers: {  
        'Accept': 'application/json;odata=nometadata',  
        'Content-type': 'application/json;odata=nometadata',  
        'odata-version': ''  
      },  
      body: body  
    })  
    .then((response: SPHttpClientResponse): Promise<IListItem> => {  
      return response.json();  
    })  
    .then((item: IListItem): void => {  
      this.setState({  
        status: `Item '${item.Title}' (ID: ${item.Id}) successfully created`,  
        Items: []  
      });  
    }, (error: any): void => {  
      this.setState({  
        status: 'Error while creating the item: ' + error,  
        Items: []  
      });  
    });  
  }   
  
  private readItem(): void {  
  }  
  
  private updateItem(): void {  
  }  
  
  private deleteItem(): void {  
  }  

  private getLatestItemId(): Promise<number> {  
    return new Promise<number>((resolve: (itemId: number) => void, reject: (error: any) => void): void => {  
      this.props.spHttpClient.get(`${this.props.siteURL}/_api/web/lists/getbytitle('${this.props.ListName}')/items?$orderby=Id desc&$top=1&$select=id`,  
        SPHttpClient.configurations.v1,  
        {  
          headers: {  
            'Accept': 'application/json;odata=nometadata',  
            'odata-version': ''  
          }  
        })  
        .then((response: SPHttpClientResponse): Promise<{ value: { Id: number }[] }> => {  
          return response.json();  
        }, (error: any): void => {  
          reject(error);  
        })  
        .then((response: { value: { Id: number }[] }): void => {  
          if (response.value.length === 0) {  
            resolve(-1);  
          }  
          else {  
            resolve(response.value[0].Id);  
          }  
        });  
    });  
  }   
  
  private getMaxId(): Promise<number>{
    return new Promise<number>((resolve: (itemId:number) => void, reject: (error: any) => void): void =>
    {
      this.props.spHttpClient.get(`${this.props.siteURL}/_api/web/lists/getbytitle('${this.props.ListName}')/items?$orderby=Id desc&top=1&$select=id`,
      SPHttpClient.configurations.v1,
      {
        headers:{
          'Accept':'application/json;odata=nometadata',
          'odataversion':''
        }
      })
      .then((response: SPHttpClientResponse): Promise<{value:{Id: number}[]}> =>{
        return response.json();
        },(error: any): void=>{
        reject(error);
      })
      .then((response: { value: { Id: number }[] }): void => {  
          if (response.value.length === 0) {  
            resolve(-1);  
          }  
          else {  
            resolve(response.value[0].Id);  
          }  
        });  
    });
  }

}
