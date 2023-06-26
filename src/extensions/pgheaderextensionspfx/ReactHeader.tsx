import * as React from "react";
// import { Link } from 'office-ui-fabric-react/lib/Link';  
// import { CommandBar } from 'office-ui-fabric-react/lib/CommandBar';  
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { Web } from '@pnp/sp/presets/all';
// import { Web } from "@pnp/sp/webs";
import { ISmarkQuickLinksPocState, IBookmarkinfo } from './ISmarkQuickLinksPocState';
export interface IReactHeaderProps {
  context: any;

}
import { parameters } from './ConstantParameters';

export default class ReactHeader extends React.Component<IReactHeaderProps, ISmarkQuickLinksPocState> {
  // private siteURLCheck:string ="";
  // private intervalID: number = 0;
  constructor(props: IReactHeaderProps) {
    super(props);
    this.state =
    {
      BookmarkDetails: undefined,
      PageExist: true,
      SiteURI: ""
    }
  }

  public async componentDidMount() {
// on-load page to check the bookmark button already enabled or not. 
    this.onCheckBookmarkButton();
  
    // this.intervalID = setInterval(() => {
    // debugger;
    //     if (this.siteURLCheck != window.location.href) {
    //       this.onCheckBookmarkButton();
          
    //       clearInterval(this.intervalID);
    //     }
    //   }, 1000);

  }
//   componentWillUnmount() { clearInterval(this.intervalID);
// }

  public onCheckBookmarkButton = async () => {

    
    const web1 = Web(parameters.webAbsURLl);
    let getListName = web1.lists.getByTitle(parameters.bookmarklistname);

    // fetching the list items which addedby "User" and Email == current login user. 
    let listItemsDeletionCheck = await getListName
      .items
      .select("Id,Title,SiteAbsoluteUrl,AddedBy,PageName,PageID,IsActive,UserEmail").filter(`AddedBy eq 'User' and UserEmail eq '${this.props.context.pageContext.user.email}'`).getAll();
    if (listItemsDeletionCheck.length > 0) {
      let checkval = 0;
      listItemsDeletionCheck.forEach(async (element, index) => {
// checking if the absoluteURL and the page id is existing ? if exist checkval will set to 1 (Disable), else it will bookmark button will get enabled.
        if (element.SiteAbsoluteUrl.toLowerCase() == this.props.context.pageContext.web.absoluteUrl.toLowerCase() && element.PageID == this.props.context.pageContext.listItem.id) {
  
          checkval = 1;
          //  this.siteURLCheck = window.location.href.toLowerCase();
          this.setState({ PageExist: true , SiteURI: window.location.href.toLowerCase() });
          
        }
      })
      // if foreach not match ? then enable the bookmark button.
      if (checkval == 0) {
        //  this.siteURLCheck = window.location.href.toLowerCase();
        this.setState({ PageExist: false , SiteURI: window.location.href.toLowerCase() });
      }
    }
    // if list item itself not found(New User), enable the bookmark.
    else {
      // this.siteURLCheck = window.location.href.toLowerCase();
      this.setState({ PageExist: false , SiteURI: window.location.href.toLowerCase() });
    }

  }
// handle event of bookmark button.
  public handleClick = () => {
    var newitem: IBookmarkinfo[] = [];
    newitem.push({
      SiteName: this.props.context.pageContext.web.title,
      PageID: this.props.context.pageContext.listItem.id,
      PageName: this.props.context.pageContext.site.serverRequestPath.split("/").slice(-1)[0],
      PageURLl: window.location.href,
      IsBookmark: "true",
      UserName: this.props.context.pageContext.user.displayName,
      UserEmail: this.props.context.pageContext.user.email,
      SiteAbsoluteUrl: this.props.context.pageContext.web.absoluteUrl
    })
    // this.siteURLCheck = window.location.href.toLowerCase();
    this.setState({ BookmarkDetails: newitem, PageExist: true }, () => {
      this.pushResult()

    }
    );

  }

  // pushing result into sharepoint bookmark list.

  private pushResult = async () => {

    const web = Web(parameters.webAbsURLl);
    await web.lists.getByTitle(parameters.bookmarklistname).items.add({

      Title: this.state.BookmarkDetails[0].SiteName,
      PageName: this.state.BookmarkDetails[0].PageName,
      PageURL: this.state.BookmarkDetails[0].PageURLl,
      IsBookmark: this.state.BookmarkDetails[0].IsBookmark,
      UserName: this.state.BookmarkDetails[0].UserName,
      UserEmail: this.state.BookmarkDetails[0].UserEmail,
      PageID: this.state.BookmarkDetails[0].PageID,
      SiteAbsoluteUrl: this.state.BookmarkDetails[0].SiteAbsoluteUrl,
      AddedBy: "User",
      IsActive: 1
    }).then(() => {

      alert("Bookmark Added Sucessfully");

    })
 
  }

  public render(): JSX.Element {
    return (
      <div style={{ float: "right", left: "-27px", top: 8 }}>
        {this.props.context.pageContext.listItem.id == undefined ?
          ""
          : <PrimaryButton style={{ float: "right", left: "-27px", top: 8 }} text="Bookmark"
            onClick={this.handleClick}
            allowDisabledFocus
            disabled={this.state.PageExist}
            checked={true} />
        }
      </div>

    );
  }


}  