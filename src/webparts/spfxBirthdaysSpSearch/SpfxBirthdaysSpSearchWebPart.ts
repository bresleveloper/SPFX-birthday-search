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
  Message: string
  Template: string;
  GetBirthdays: string;
  //BGcolor: string;
  //FontColor: string;
  AddShadow: boolean;
  Debug: boolean;
}

export default class SpfxBirthdaysSpSearchWebPart extends BaseClientSideWebPart<ISpfxBirthdaysSpSearchWebPartProps> {

  public templates = {
    '1 line no image': ` 
      <a class="${styles.flex} ${styles.lineNoImage} ${styles.cleanA}" href="#MAILTO#">
        <span class="${styles.preffix}">#PREFFIX#</span>
        <span class="${styles.date}">#DATE#</span>
        <span class="name">#NAME#</span>
        <span class="${styles.suffix}">#SUFFIX#</span>
      </a>
    `,
    '3 lines with image': ` 
      <a class="${styles.lineWithImage} ${styles.flex} ${styles.cleanA}" href="#MAILTO#">
        <div class="img">#IMG#</div>

        <div class="${styles.details} ${styles["flex-col"]}">
          <span class="preffix">#PREFFIX#</span>
          <span class="date">#DATE#</span>
          <span class="name">#NAME#</span>
          <span class="suffix">#SUFFIX#</span>
        </div>
      </a>
    `,
    'image-title-department': ` 
    <a class="${styles.lineWithImage} ${styles.flex} ${styles.cleanA}" href="#MAILTO#">
      <div class="${styles.lineWithImage}">#IMG#</div>

      <div class="${styles.details} ${styles["flex-col"]}">
        <div>
          <span class="preffix">#PREFFIX#</span>
          <span class="${styles.name} ${styles.big}">#NAME#</span>
          <span class="suffix">#SUFFIX#</span>
        </div>
        <span class="${styles.department}">#DEP#</span>
      </div>
    </a>
  `,
    '3-lines-image-dark': `

      <a class="lines-image-dark ${styles.lineImageDark} ${styles.flex} ${styles.cleanA}" href="#MAILTO#">
      <div class="${styles.b}">#IMG#</div>
      <div class=" ${styles["flex-col"]}">
        <div>
          <span class="${styles.preffix}">#PREFFIX#</span>
          <span class="${styles.name} ${styles.a}">#FIRSTNAME#</span>
          <br/>
          <span class="${styles.date}">#DATE#</span>
          <span class="${styles.name}">#NAME#</span>
          <span class="${styles.suffix}">#SUFFIX#</span>
          <button class="${styles}">#MAILTO#</button>
        </div>
        <span class="${styles.department}">#DEP#</span>
      </div>
    </a>
    `
  }

  public render(): void {
    this.properties.Debug = true;
    //console.log('I R THIS', this)
    this.domElement.innerHTML = `<h2>Loading Birthdays</h2>`
    if (this.properties.Debug == true) {
      this.buildHtml([{
        // Culture: "he-IL",
        //DocId: "17656901765554",
        // DocumentSignature: "",
        //EditProfileUrl: null,
        //GeoLocationSource: "EUR",
        //IsExternalContent: "false",
        //ListId: null,
        //PartitionId: "4b4fc818-94f8-44a2-9541-4cea3e234001",
        PictureURL: "https://publicdomainvectors.org/tn_img/eco-systemedic-star.png",
        PreferredName: "יניב גולדגלס",
        //ProfileQueriesFoundYou: null,
        // ProfileViewsLastMonth: null,
        //ProfileViewsLastWeek: null,
        //Rank: "16.9419651",
        RefinableString99: "5/20/2000 12:00:00 AM",
        //RenderTemplateId: "~sitecollection/_catalogs/masterpage/Display Templates/Search/Item_Default.js",
        //ResultTypeId: null,
        //ResultTypeIdList: null,
        // SiteId: null,
        Title: "יניב גולדגלס",
        //UniqueId: null,
        //: "0",
        //WebId: null,
        WorkEmail: "tehila1728@gmail.com",
        // contentclass: "urn:content-class:SPSPeople",
        date: new Date('Sat May 20 2000 00:00:00 GMT+0300 (Israel Daylight Time)'),
        // piSearchResultId: "ARIAǂ1638d3bf-7f33-4e9a-a1d7-02779662f857ǂ0c7ee126-4faa-4b30-adc5-2a0caa8f1c4cǂ1638d3bf-7f33-4e9a-a1d7-02779662f857.1000.1ǂEUR:unknown:",
        showDateStr: "20.4",
        Message: "שלח ברכה"
      }, {
        PictureURL: 'https://publicdomainvectors.org/tn_img/five_pointed_star.png',
        FirstName: 'refaeli',
        PreferredName: 'תהילה',
        RefinableString99: "5/20/2000 12:00:00 AM",
        Title: 'refaeli',
        date: new Date('Sat May 20 2000 00:00:00 GMT+0300 (Israel Daylight Time)'),
        showDateStr: "20.4",
        WorkEmail: "tehila1728@gmail.com",
        Message: "שליחת ברכה"



      }])

      return;
    }
    this.getBirthdays();
  }

  public getBirthdays() {
    let currentMonth: any = new Date().getMonth() + 1
    let nextMonth = currentMonth + 1;
    nextMonth = nextMonth == 13 ? 1 : nextMonth
    //was wrong...
    //currentMonth = currentMonth >= 10 ? currentMonth : '0' + currentMonth
    //nextMonth = nextMonth >= 10 ? nextMonth : '0' + nextMonth

    //changed Birthday to RefinableString99
    let searchQ = "querytext='RefinableString99:" + currentMonth +
      " OR RefinableString99:" + nextMonth +
      "'&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'" +
      "&rowlimit=1000&selectproperties='Title,WorkEmail,PreferredName,FirstName,PictureURL,RefinableString99'"

    if (this.properties.GetBirthdays && this.properties.GetBirthdays == "Today") {
      let day: any = new Date().getDate()
      day = day >= 10 ? day : '0' + day;

      searchQ = "querytext='RefinableString99:" + currentMonth +
        " AND RefinableString99:" + day +
        "'&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'" +
        "&rowlimit=1000&selectproperties='Title,WorkEmail,PreferredName,FirstName,PictureURL," +
        "RefinableString99,RefinableString98,RefinableString97,Department'"
    }

    //debug
    //searchQ = "querytext='Birthday:8'&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'" +
    //searchQ = "querytext='RefinableString99:8'&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'" +
    //"&rowlimit=1000&selectproperties='Title,WorkEmail,PreferredName,PictureURL,Birthday'"
    //"&rowlimit=1000&selectproperties='Title,WorkEmail,PreferredName,PictureURL,RefinableString99'"

    this.search(searchQ, (arr) => {
      console.log('arr in search callback', arr);
      let arr2 = []
      let currentMonth: any = new Date().getMonth() + 1
      let nextMonth = currentMonth + 1;
      nextMonth = nextMonth == 13 ? 1 : nextMonth
      //was wrong...
      //currentMonth = currentMonth >= 10 ? currentMonth : '0' + currentMonth
      //nextMonth = nextMonth >= 10 ? nextMonth : '0' + nextMonth

      //for some reason there are dups, i'll trim by email, which can be null
      let emailNamesKeys = {}

      //debugger
      //22/08/2000 00:00:00
      for (let i = 0; i < arr.length; i++) {
        const up = arr[i];
        let key = up['WorkEmail'] ? up['WorkEmail'] : up['PreferredName']

        if (emailNamesKeys[key]) {
          continue
        } else {
          emailNamesKeys[key] = true
        }//ohh common update the damn file

        //changed Birthday to RefinableString99
        //RefinableString99: "1/18/2000 12:00:00 AM"
        let dArr = up.RefinableString99.split(" ")[0].split("/")
        let todayDay = new Date().getDate()
        //let bDay = parseInt(dArr[0])
        console.log('dArr', dArr);

        let bDay = parseInt(dArr[1])
        let bMonth = parseInt(dArr[0])

        if (this.properties.GetBirthdays && this.properties.GetBirthdays == "Today") {
          if (bMonth == currentMonth && bDay == todayDay) {
            //up.showDateStr = dArr[1] + '.' + dArr[0]
            //up.date = new Date(2000, dArr[1], dArr[0])
            up.showDateStr = bDay + '.' + currentMonth
            up.date = new Date(2000, currentMonth, bDay)
            arr2.push(up)
          }
        } else if (    //true || //debug // month foreward
          //(dArr[1] == currentMonth && bDay >= todayDay) ||
          //(dArr[1] == nextMonth && bDay <= todayDay) 
          (currentMonth == currentMonth && bDay >= todayDay) ||
          (currentMonth == nextMonth && bDay <= todayDay)
        ) {
          //up.showDateStr = dArr[1] + '.' + dArr[0]
          //up.date = new Date(2000, dArr[1], dArr[0])
          up.showDateStr = bDay + '.' + currentMonth
          up.date = new Date(2000, currentMonth, bDay)
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

  public buildHtml(arr2) {
    console.log(this.context.pageContext.user)

    let myName = this.context.pageContext.user.displayName;

    let h2 = '';
    h2 = `<div class="${styles.spfxBirthdaysSpSearch}">
            <div class="#SPECIAL# ${styles["flex-col"]} ${this.properties.AddShadow ? styles.shadowWrapper : styles.justPad}">
              <h2>${this.properties.Title ? this.properties.Title : 'ימי הולדת'}</h2>
              <div class="${styles["flex-col"]}">`

    let t = this.properties.Template ? this.templates[this.properties.Template] : this.templates['1 line no image'];
    for (let i = 0; i < arr2.length; i++) {
      const x = arr2[i];
      let yourName = x['PreferredName'] ? x['PreferredName'] : x['Title']
      let firstName = x['FirstName'] ? x['FirstName'] : x['']

      h2 += t.replace('#MAILTO#', `mailto:${x.WorkEmail}?subject=Happy Birthday From ${myName}`)
        .replace('#IMG#', (x.PictureURL ? "<img src=\"" + x.PictureURL + "\">" : ''))
        .replace('#PREFFIX#', this.properties.Preffix ? this.properties.Preffix : '')
        .replace('#DATE#', x.showDateStr)
        .replace('#NAME#', yourName)
        .replace('#DEP#', (x.RefinableString97 ? x.RefinableString97 : ''))
        .replace('#SUFFIX#', this.properties.Suffix ? this.properties.Suffix : '')
        .replace('#MAILTO#', this.properties.Message ? this.properties.Message : 'שליחת ברכה')
        .replace('#FIRSTNAME#', firstName)
    }

    h2 += '</div></div></div>'
    this.domElement.innerHTML = h2;

    let h3: string;
    let three_line_dark = document.querySelector('.lines-image-dark');
    console.log("three_line_dark ", three_line_dark);

    let is_three_line_dark = false;
    if (three_line_dark != null) {
      is_three_line_dark = true
    }
    if (is_three_line_dark == true) {
      console.log("the template is dark");

      h3 = h2.replace('#SPECIAL#', `${styles["lines-image-dark"]}`)
    }
    else {
      h3 = h2
    }

    this.domElement.innerHTML = h3;
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

  public search(query: string, callback): void {

    console.log('search query', query);
    //this.ajaxCounter++;

    this.context.spHttpClient.get(
      this.context.pageContext.web.absoluteUrl +
      `/_api/search/query?` + query, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        response.json().then((data) => {

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
                PropertyPaneTextField('Title', { label: 'Title' }),
                PropertyPaneTextField('Preffix', { label: 'Preffix' }),
                PropertyPaneTextField('Suffix', { label: 'Suffix' }),
                //https://techcommunity.microsoft.com/t5/sharepoint-developer/propertypanecheckbox-default-state-issue/m-p/75946
                //PropertyPaneCheckbox('Template', {})
                PropertyPaneDropdown('Template', {
                  label: 'Template',
                  options: [
                    { key: 'image-title-department', text: 'image-title-department' },
                    { key: '1 line no image', text: '1 line no image' },
                    { key: '3 lines with image', text: '3 lines with image' },
                    { key: '3-lines-image-dark', text: '3-lines-image-dark' },
                  ]
                }),
                PropertyPaneDropdown('GetBirthdays', {
                  label: 'Get Birthdays',
                  options: [
                    { key: 'Month Forward', text: 'This Forward' },
                    { key: 'Today', text: 'Today' },
                  ]
                }),
                //PropertyPaneTextField('BGcolor', {label:'Background Color'}),
                //PropertyPaneTextField('FontColor', {label:'Font Color'}),
                PropertyPaneCheckbox('AddShadow', { text: 'Add Shadow Box' }),
                PropertyPaneCheckbox('Debug', { text: 'Debug' }),

              ]//end groupFields
            }
          ]
        }
      ]
    };
  }
}
