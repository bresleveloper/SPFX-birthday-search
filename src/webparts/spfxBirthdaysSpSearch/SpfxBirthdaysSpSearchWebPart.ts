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
    /*this.domElement.innerHTML = `
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
      </div>`;*/
      this.domElement.innerHTML = `<h2>Loading Birthdays</h2>`

      this.getBirthdays();
  } 


  public getBirthdays(){
    let currentMonth:any = new Date().getMonth() + 1
    let nextMonth = currentMonth +1;
    nextMonth = nextMonth == 13 ? 1 : nextMonth
    currentMonth = currentMonth >= 10 ? currentMonth : '0' + currentMonth
    nextMonth = nextMonth >= 10 ? nextMonth : '0' + nextMonth
    
    let searchQ = "querytext='Birthday:" + currentMonth + 
        " OR Birthday:" + nextMonth + 
        "'&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'" +
        "&rowlimit=1000&selectproperties='Title,WorkEmail,PreferredName,PictureURL,Birthday'"

    //debug
    searchQ = "querytext='Birthday:8'&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'" +
    "&rowlimit=1000&selectproperties='Title,WorkEmail,PreferredName,PictureURL,Birthday'"

    this.search(searchQ, (arr) => {
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

          let dArr = up.Birthday.split(" ")[0].split("/")
          let todayDay = new Date().getDate()
          let bDay = parseInt(dArr[0])

          if (    true || //debug
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
      this.buildHtml(arr2)
    });
  }

  public buildHtml(arr2){
    console.log(this.context.pageContext.user)
  
    let myName = this.context.pageContext.user.displayName;
    let h2 = '';
    h2 = '<div class="mdl-tabs__panel" >'
    //'Title,WorkEmail,PreferredName,PictureURL,Birthday'
    arr2.forEach(function (b) {
      h2 += "<div class=\"article\">" +
              "<div class=\"date light_blue fs16 inline-block\">" + b.showDateStr + "</div>" +
              "<div class=\"excerptContainer inline-block\">" +
                "<div class=\"excerpt dark_grey fs18\">" +
                  b.PreferredName + " - " +
                  "<a href=\"mailto:" + b.WorkEmail + "?subject=Happy Birthday From " + myName + "\">שלח ברכה</a>" +
                "</div>" +
                (b.PictureURL ? "<img src=\"" + b.PictureURL + "\">" : '') +
              "</div>" +
            "</div>"
    });
    h2 += '</div>'

    this.domElement.innerHTML = h2;
  }

  public searchOld(querystring, callback) {
    try {
        let reqListenerSearchParser = function reqListenerSearchParser() {
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
                  let results = data.PrimaryQueryResult.RelevantResults.Table.Rows;
                  let arr = []
  
                  console.log('rows', results);

                  results.forEach(function (row) {
                      let item = {}
                      row.Cells.forEach(function (cell) { item[cell.Key] = cell.Value })
                      arr.push(item)
                  });

                  console.log('normalized', arr);
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
