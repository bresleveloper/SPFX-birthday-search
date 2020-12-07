import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SpfxBirthdaysSpSearchWebPart.module.scss';
import * as strings from 'SpfxBirthdaysSpSearchWebPartStrings';



import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';


export interface ISpfxBirthdaysSpSearchWebPartProps {
  description: string;
}

export default class SpfxBirthdaysSpSearchWebPart extends BaseClientSideWebPart<ISpfxBirthdaysSpSearchWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.spfxBirthdaysSpSearch }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
  } 


  public getBirthdays(){
    let currentMonth:any = new Date().getMonth() + 1
    let nextMonth = currentMonth +1;
    nextMonth = nextMonth == 13 ? 1 : nextMonth
    currentMonth = currentMonth >= 10 ? currentMonth : '0' + currentMonth
    nextMonth = nextMonth >= 10 ? nextMonth : '0' + nextMonth
    
    let searchQ = "querytext='BirthdayString:" + currentMonth + 
        " OR BirthdayString:" + nextMonth + 
        "'&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'" +
        "&rowlimit=1000&selectproperties='Title,WorkEmail,PreferredName,PictureURL,BirthdayString'"

    this.search(searchQ, function (arr) {
      let arr2 = []
      let currentMonth:any = new Date().getMonth() + 1
      let nextMonth = currentMonth +1;
      nextMonth = nextMonth == 13 ? 1 : nextMonth
      currentMonth = currentMonth >= 10 ? currentMonth : '0' + currentMonth
      nextMonth = nextMonth >= 10 ? nextMonth : '0' + nextMonth

      //for some reason there are dups, i'll trim by email, which can be null
      let emailNamesKeys = {}

      //22/08/2000 00:00:00
      for (let i = 0; i < arr.length; i++) {
          const up = arr[i];
          let key = up['WorkEmail'] ? up['WorkEmail'] : up['PreferredName']

          if (emailNamesKeys[key]) {
              continue
          } else {
              emailNamesKeys[key] = true
          }//ohh common update the damn file

          let dArr = up.BirthdayString.split(" ")[0].split("/")
          let todayDay = new Date().getDate()
          let bDay = parseInt(dArr[0])

          if ( 
                  (dArr[1] == currentMonth && bDay >= todayDay) ||
                  (dArr[1] == nextMonth && bDay <= todayDay) 
              ) {
              up.showDateStr = dArr[0] + '.' + dArr[1]
              up.date = new Date(2000, dArr[1], dArr[0])
              arr2.push(up)
          } 
      }

      arr2.sort(function (a, b) {
          return a.date < b.date ? -1 : 1
      })

      console.log('birthdays arr', arr2);
      
      //comps.events.birthdaysArr = arr2
      //call a fn that will build some html
    });
  }


  public searchOld(querystring, callback) {
    try {
        function reqListenerSearchParser() {
            try {
                let searchResultsFull = JSON.parse(this.responseText)
                let results = searchResultsFull.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results;
                let arr = []

                results.forEach(function (row) {
                    let item = {}
                    row.Cells.results.forEach(function (cell) { item[cell.Key] = cell.Value })
                    arr.push(item)
                });

                callback(arr);
            } catch (error) {
                console.error('search error (reqListenerSearchParser)')
                console.error(error)
                callback(null)
            }
        }

        let oReq = new XMLHttpRequest();
        oReq.addEventListener("load", reqListenerSearchParser);
        oReq.open("GET", this.context.pageContext.web.absoluteUrl + "/_api/search/query?" + querystring);
        oReq.setRequestHeader("Accept", "application/json;odata=verbose");
        oReq.send();
    } catch (e) {
        console.error('search error')
        console.error(e)
        callback(null)
    }
  }

  public search(query:string, callback): void {

    console.log('search query', query);
    //this.ajaxCounter++;

    this.context.spHttpClient.get(
      this.context.pageContext.web.absoluteUrl +
      `/_api/search/query?` + query, SPHttpClient.configurations.v1)
          .then((response: SPHttpClientResponse) => {
              response.json().then((data)=> {

                  console.log('search results', query, data);
                  //this.listsContainer[listname] = data.value;

                  //assuming that value is the search res standard
                  //let searchResultsFull = JSON.parse(this.responseText)
                  let results = data.value.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results;
                  let arr = []
  
                  results.forEach(function (row) {
                      let item = {}
                      row.Cells.results.forEach(function (cell) { item[cell.Key] = cell.Value })
                      arr.push(item)
                  });


                  callback(arr)
              });
          });
    }


  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
