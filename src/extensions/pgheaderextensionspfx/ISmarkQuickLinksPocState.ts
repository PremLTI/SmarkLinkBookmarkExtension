export interface IBookmarkinfo{
    SiteName: string;
    PageName: string ;
    PageURLl: string;
    IsBookmark: string;
    UserName:string;
    UserEmail: string;
    PageID:string;
    SiteAbsoluteUrl:string;
   
    
}

export interface ISmarkQuickLinksPocState {

    BookmarkDetails: IBookmarkinfo[] ;
    PageExist:boolean;
    SiteURI: string;

  }