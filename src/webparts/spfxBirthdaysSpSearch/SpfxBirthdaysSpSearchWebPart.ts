import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
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
  Title: string;
  Preffix: string;
  Suffix: string;
  Template: string;
  GetBirthdays: string;
}

export default class SpfxBirthdaysSpSearchWebPart extends BaseClientSideWebPart<ISpfxBirthdaysSpSearchWebPartProps> {

  public templates = {
    '1 line no image' : ` 
      <a class="${styles.flex} ${styles.lineNoImage}" href="#MAILTO#">
        <span class="${styles.preffix}">#PREFFIX#</span>
        <span class="${styles.date}">#DATE#</span>
        <span class="name">#NAME#</span>
        <span class="${styles.suffix}">#SUFFIX#</span>
      </a>
    `,
    '3 lines with image' : ` 
      <a class="lines-with-img-item ${styles.flex}" href="#MAILTO#">
        <div class="img">
          <img src="#SRC#"/>
        </div>

        <div class="details ${styles["flex-col"]}">
          <span class="preffix">#PREFFIX#"</span>
          <span class="date">#DATE#"</span>
          <span class="name">#NAME#"</span>
          <span class="suffix">#SUFFIX#"</span>
        </div>
      </a>
    `
  }

  public render(): void {
    //console.log('I R THIS', this)
    this.domElement.innerHTML = `<h2>Loading Birthdays</h2>`
    this.getBirthdays();
  } 

  public getBirthdays(){
    let currentMonth:any = new Date().getMonth() + 1
    let nextMonth = currentMonth +1;
    nextMonth = nextMonth == 13 ? 1 : nextMonth
    currentMonth = currentMonth >= 10 ? currentMonth : '0' + currentMonth
    nextMonth = nextMonth >= 10 ? nextMonth : '0' + nextMonth
    
    //changed Birthday to RefinableDate07
    let searchQ = "querytext='RefinableDate07:" + currentMonth + 
        " OR RefinableDate07:" + nextMonth + 
        "'&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'" +
        "&rowlimit=1000&selectproperties='Title,WorkEmail,PreferredName,PictureURL,RefinableDate07'"

    if (this.properties.GetBirthdays && this.properties.GetBirthdays == "Today") {
      let day:any = new Date().getDate()
      day = day >= 10 ? day : '0' + day;

      searchQ = "querytext='RefinableDate07:" + currentMonth + 
        " AND RefinableDate07:" + day + 
        "'&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'" +
        "&rowlimit=1000&selectproperties='Title,WorkEmail,PreferredName,PictureURL,RefinableDate07'"
    }

    //debug
    //searchQ = "querytext='Birthday:8'&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'" +
    //searchQ = "querytext='RefinableDate07:8'&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'" +
    //"&rowlimit=1000&selectproperties='Title,WorkEmail,PreferredName,PictureURL,Birthday'"
    //"&rowlimit=1000&selectproperties='Title,WorkEmail,PreferredName,PictureURL,RefinableDate07'"

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

          //changed Birthday to RefinableDate07
          let dArr = up.RefinableDate07.split(" ")[0].split("/")
          let todayDay = new Date().getDate()
          let bDay = parseInt(dArr[0])

          if (this.properties.GetBirthdays && this.properties.GetBirthdays == "Today") {
            if (dArr[1] == currentMonth && bDay == todayDay) {
              up.showDateStr = dArr[1] + '.' + dArr[0]
              up.date = new Date(2000, dArr[1], dArr[0])
              arr2.push(up)
            }
          } else if (    //true || //debug // month foreward
                  (dArr[1] == currentMonth && bDay >= todayDay) ||
                  (dArr[1] == nextMonth && bDay <= todayDay) 
              ) {
              //up.showDateStr = dArr[0] + '.' + dArr[1]
              up.showDateStr = dArr[1] + '.' + dArr[0]
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
    h2 = `<div class="${styles.spfxBirthdaysSpSearch}">
            <div class="${styles["flex-col"]} ${styles.shadowWrapper}">
              <h2>${this.properties.Title ? this.properties.Title : 'ימי הולדת'}</h2>
              <div class="${styles["flex-col"]}">`

    let t = this.properties.Template ? this.templates[this.properties.Template] : this.templates['1 line no image'];
    for (let i = 0; i < arr2.length; i++) {
      const x = arr2[i];
      h2 += t.replace('#MAILTO#', `mailto:${x.WorkEmail}?subject=Happy Birthday From ${myName}`)
              .replace('#SRC#', (x.PictureURL ? "<img src=\"" + x.PictureURL + "\">" : ''))
              .replace('#PREFFIX#',this.properties.Preffix ? this.properties.Preffix : '')
              .replace('#DATE#', x.showDateStr)
              .replace('#NAME#', myName)
              .replace('#SUFFIX#',this.properties.Suffix ? this.properties.Suffix : '')
    }

    h2 += '</div></div></div>'

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
                PropertyPaneTextField('Title', {label: 'Title'}),
                PropertyPaneTextField('Preffix', {label:'Preffix'}),
                PropertyPaneTextField('Suffix', {label:'Suffix'}),
                //https://techcommunity.microsoft.com/t5/sharepoint-developer/propertypanecheckbox-default-state-issue/m-p/75946
                //PropertyPaneCheckbox('Template', {})
                PropertyPaneDropdown('Template', {label:'Template', 
                  options:[
                    {key:'1 line no image',text:'1 line no image'},
                    {key:'3 lines with image',text:'3 lines with image'},
                  ]
                }),
                PropertyPaneDropdown('GetBirthdays', {label:'Get Birthdays', 
                  options:[
                    {key:'Month Forward',text:'This Forward'},
                    {key:'Today',text:'Today'},
                  ]
                }),
              ]//end groupFields
            }
          ]
        }
      ]
    };
  }
}
