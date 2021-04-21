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
  bringAllBirthdays: boolean;
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
        <div class="${styles.img}">#IMG#</div>

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
      <div class="${styles.img}">#IMG#</div>

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
      <div class="${styles.flex}">
        <div class="${styles.img}">#IMG#</div>
        <div class=" ${styles["flex-col"]}">
          <div class="">
              <span class="${styles.preffix}">#PREFFIX#</span>
              <span class="${styles.name} ${styles.boldA}">#NAME#</span>
              <span class="${styles.suffix}">#SUFFIX#</span>
              <span class="date">#DATE#</span>
          </div>
          <span class="${styles.department}">#DEP#</span>
        </div>
      </div>
      <div class="${styles['send-bless']}">
        <div class="${styles.darkButton}">שליחת ברכה</div>
      </div>
    </a>`
  }

  public render(): void {
    this.domElement.innerHTML = `<h2>Loading Birthdays</h2>`
    if (this.properties.Debug == true) {
      this.buildHtml([{
        PictureURL: "https://images.squarespace-cdn.com/content/v1/5a7c0544d74cffa3a6ce66b3/1587740850248-HT0QC4V60Y17PK8D1A6F/ke17ZwdGBToddI8pDm48kFWxnDtCdRm2WA9rXcwtIYR7gQa3H78H3Y0txjaiv_0fDoOvxcdMmMKkDsyUqMSsMWxHk725yiiHCCLfrh8O1z5QPOohDIaIeljMHgDF5CVlOqpeNLcJ80NK65_fV7S1UcTSrQkGwCGRqSxozz07hWZrYGYYH8sg4qn8Lpf9k1pYMHPsat2_S1jaQY3SwdyaXg/%D7%AA%D7%9E%D7%95%D7%A0%D7%AA+%D7%A0%D7%95%D7%A3+-+%D7%90%D7%92%D7%9D++%D7%92%D7%90%D7%A8%D7%93%D7%94+%D7%90%D7%99%D7%98%D7%9C%D7%99%D7%94.jpg?format=2500w",
        PreferredName: "רונית שרייבר",
        RefinableString99: "5/20/2000 12:00:00 AM",
        RefinableString97: "קיסריות",
        Title: "רונית שרייבר",
        WorkEmail: "tehila1728@gmail.com",
        date: new Date('Sat May 20 2000 00:00:00 GMT+0300 (Israel Daylight Time)'),
        showDateStr: "20.4",
      }, {
        PictureURL: 'https://www.sananes.co.il/media/catalog/product/cache/1/thumbnail/795x/17f82f742ffe127f42dca9de82fb58b1/0/1/018ve.jpg',
        FirstName: 'refaeli',
        PreferredName: 'תהילה רפאלי',
        RefinableString99: "5/20/2000 12:00:00 AM",
        RefinableString97: "מתארסות",
        Title: 'refaeli',
        date: new Date('Sat May 20 2000 00:00:00 GMT+0300 (Israel Daylight Time)'),
        showDateStr: "21.4",
        WorkEmail: "tehila1728@gmail.com",
      }, {
        PictureURL: 'https://www.sananes.co.il/media/catalog/product/cache/1/thumbnail/795x/17f82f742ffe127f42dca9de82fb58b1/0/1/018ve.jpg',
        FirstName: 'refaeli',
        PreferredName: 'תהילה רפאלי',
        RefinableString99: "5/20/2000 12:00:00 AM",
        RefinableString97: "מתארסות",
        Title: 'refaeli',
        date: new Date('Sat May 20 2000 00:00:00 GMT+0300 (Israel Daylight Time)'),
        showDateStr: "22.4",
        WorkEmail: "tehila1728@gmail.com",
      }, {
        PictureURL: 'https://www.sananes.co.il/media/catalog/product/cache/1/thumbnail/795x/17f82f742ffe127f42dca9de82fb58b1/0/1/018ve.jpg',
        FirstName: 'refaeli',
        PreferredName: 'תהילה רפאלי',
        RefinableString99: "5/20/2000 12:00:00 AM",
        RefinableString97: "מתארסות",
        Title: 'refaeli',
        date: new Date('Sat May 20 2000 00:00:00 GMT+0300 (Israel Daylight Time)'),
        showDateStr: "20.4",
        WorkEmail: "tehila1728@gmail.com",
      }, {
        PictureURL: 'https://www.sananes.co.il/media/catalog/product/cache/1/thumbnail/795x/17f82f742ffe127f42dca9de82fb58b1/0/1/018ve.jpg',
        FirstName: 'refaeli',
        PreferredName: 'תהילה רפאלי',
        RefinableString99: "5/20/2000 12:00:00 AM",
        RefinableString97: "מתארסות",
        Title: 'refaeli',
        date: new Date('Sat May 20 2000 00:00:00 GMT+0300 (Israel Daylight Time)'),
        showDateStr: "20.4",
        WorkEmail: "tehila1728@gmail.com",
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
    let end = "'&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'" +
      "&rowlimit=1000&selectproperties='Title,WorkEmail,PreferredName,FirstName,PictureURL," +
      "RefinableString99,RefinableString98,RefinableString97,RefinableString95,Department,Birthday'"

    let searchQ = "querytext='RefinableString99:" + currentMonth +
      " OR RefinableString99:" + nextMonth + end

    if (this.properties.GetBirthdays && this.properties.GetBirthdays == "Today") {
      let day: any = new Date().getDate()
      day = day >= 10 ? day : '0' + day;

      searchQ = "querytext='RefinableString99:" + currentMonth +
        " AND RefinableString99:" + day + end
    }

    if (this.properties.bringAllBirthdays == true) {
      searchQ = "querytext='RefinableString99:1900" + end
    }

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
        up.showDateStr = bDay + '.' + currentMonth

        if (true) {
          let d = bDay < 10 ? "0" + bDay : bDay;
          let m = currentMonth < 10 ? "0" + currentMonth : currentMonth;
          up.showDateStr = d + '/' + m  
        }

        up.date = new Date(2000, currentMonth, bDay)
        if (this.properties.bringAllBirthdays == true) {
          arr2.push(up)
          continue
        }

        if (this.properties.GetBirthdays && this.properties.GetBirthdays == "Today") {
          if (bMonth == currentMonth && bDay == todayDay) {
            //up.showDateStr = dArr[1] + '.' + dArr[0]
            //up.date = new Date(2000, dArr[1], dArr[0])
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
    let shadowClass = this.properties.AddShadow ? styles.shadowWrapper : styles.justPad
    let noBG = false
    if (this.properties.Template == '3-lines-image-dark') {
      shadowClass = styles.lineImageDark;
      noBG = true
    }
    h2 = `<div class="${styles.spfxBirthdaysSpSearch} ${noBG ? styles.noBG : ""}">
            <div class="${styles["flex-col"]} ${shadowClass}">
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
                PropertyPaneCheckbox('bringAllBirthdays', { text: 'Bring All Birthdays' }),

              ]//end groupFields
            }
          ]
        }
      ]
    };
  }
}
