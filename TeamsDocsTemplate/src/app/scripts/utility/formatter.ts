import { IBreadCrumb } from "../interfaces/IBreadCrumb";
import { IFileFolder } from "../interfaces/IFileFolder";
import { DEFAULT_XLSX_URL, DEFAULT_PPTX_URL, DEFAULT_DOCX_URL, TEMPLATE_IMAGE_PATH, DEFAULT_FOLDER_URL, DENIED_FOLDER_URL } from "./Constants";
import { IBreadcrumb, IBreadCrumbData } from "office-ui-fabric-react";

export function USformatDate(date) {
    let inDate = new Date(date);
    return inDate.getMonth() + 1 + "/" + inDate.getDate() + "/" + inDate.getFullYear();
}

export function compare(a, b) {
    // Use toUpperCase() to ignore character casing
    const templateA = a.toShow;
    const templateB = b.toShow;

    let comparison = 0;
    if (templateA > templateB) {
        comparison = -1;
    } else if (templateA < templateB) {
        comparison = 1;
    }
    return comparison;
}

export function loadBreadCrumb(Location: string, currentBreadCrumb: IBreadCrumb[], ResponseData: any, Position: number, LocationDisplay: string): IBreadCrumb[] {

    let finalBreadCrumb: IBreadCrumb[] = [...currentBreadCrumb];

    let indexOflocation = finalBreadCrumb.map(data => data.Location).indexOf(Location);


    let finalContents: IFileFolder[] = [];

    if (ResponseData.Files) {
        ResponseData.Files.results.map((data) => {
            let currentContent: IFileFolder = {
                ParentLocation: Location,
                Name: data.Name,
                ServerRelativeUrl: data.ServerRelativeUrl,
                Title: data.Title,
                UniqueId: data.UniqueId,
                type: data.__metadata.type,
                TimeLastModified: data.TimeLastModified,
                HasAccess: true
            };

            finalContents.push(currentContent);
        });
    }
    if (ResponseData.Folders) {
        ResponseData.Folders.results.map((data) => {
            if (data.Name.toUpperCase() !== "FORMS") {
                let currentContent: IFileFolder = {
                    ParentLocation: Location,
                    Name: data.Name,
                    ServerRelativeUrl: data.ServerRelativeUrl,
                    Title: data.Title,
                    UniqueId: data.UniqueId,
                    type: data.__metadata.type,
                    TimeLastModified: data.TimeLastModified,
                    HasAccess: true
                };

                finalContents.push(currentContent);
            }
        });

    }

    if (indexOflocation > -1) {
        finalBreadCrumb[indexOflocation].Contents = finalContents;
        finalBreadCrumb[indexOflocation].Order = Position;
        finalBreadCrumb[indexOflocation].InCurrentView = true;
        finalBreadCrumb[indexOflocation].IsSkipped = false;
        finalBreadCrumb[indexOflocation].Location = Location;
        finalBreadCrumb[indexOflocation].LocationDisplay = LocationDisplay ? LocationDisplay : "Home";
    }

    else {
        let currentItem: IBreadCrumb = {
            Contents: finalContents,
            IsSkipped: false,
            Order: Position,
            InCurrentView: true,
            Location: Location,
            LocationDisplay: LocationDisplay ? LocationDisplay : "Home"
        };
        finalBreadCrumb.push(currentItem);
    }

    finalBreadCrumb = showHideBreadCrumb(finalBreadCrumb, Location, Position);

    return finalBreadCrumb.sort(sortBreadCrumb);

}

export function sortBreadCrumb(firstBreadCrumb: IBreadCrumb, secondBreadCrumb: IBreadCrumb) {
    // Use toUpperCase() to ignore character casing
    const templateA = firstBreadCrumb.Order;
    const templateB = secondBreadCrumb.Order;

    let comparison = 0;
    if (templateA > templateB) {
        comparison = -1;
    } else if (templateA < templateB) {
        comparison = 1;
    }
    return comparison;
}

export function showHideBreadCrumb(finalBreadCrumb: IBreadCrumb[], Location: string, Position: number): IBreadCrumb[] {
    // Use toUpperCase() to ignore character casing

    let ChangedBreadCrumbs: IBreadCrumb[] = [];

    finalBreadCrumb.map(
        (data: IBreadCrumb) => {
            let ChangedBreadCrumb: IBreadCrumb = { ...data };
            if (ChangedBreadCrumb.Order >= Position && ChangedBreadCrumb.Location !== Location) {
                ChangedBreadCrumb.InCurrentView = false;
            }
            else if (ChangedBreadCrumb.Location === Location) {
                ChangedBreadCrumb.InCurrentView = true;
            }
            ChangedBreadCrumbs.push(ChangedBreadCrumb);
        }
    );

    return ChangedBreadCrumbs;

}

export function updateBreadCrumbForAccessDenied(finalBreadCrumb: IBreadCrumb[], UniqueId: string): IBreadCrumb[] {
    // Use toUpperCase() to ignore character casing

    let ChangedBreadCrumbs: IBreadCrumb[] = [];

    finalBreadCrumb.map(
        (data: IBreadCrumb) => {

            let indexOfDeniedContent = data.Contents.map((d: IFileFolder) => d.UniqueId).indexOf(UniqueId);
            if (indexOfDeniedContent > -1) {
                data.Contents[indexOfDeniedContent].HasAccess = false;
            }
            ChangedBreadCrumbs.push(data);
        }
    );

    return ChangedBreadCrumbs;

}

export function populateTemplates(context: any, fileFolderList: IFileFolder[]): any[] {

    let currentTemplates: any[] = [];

    fileFolderList.map(template => {
        let eachTemplate: any = { ...template };
        eachTemplate.isSelected = false;
        eachTemplate.toShow = true;
        if (template.ServerRelativeUrl) {
            if (template.type === "SP.Folder") {
                eachTemplate.imgUrl = template.HasAccess ? DEFAULT_FOLDER_URL : DENIED_FOLDER_URL;
                eachTemplate.altImgUrl = template.HasAccess ? DEFAULT_FOLDER_URL : DENIED_FOLDER_URL;
            }
            else {
                eachTemplate.imgUrl = TEMPLATE_IMAGE_PATH + `&tenant=${context.tid}` +
                    `&SPOUrl=` + `https://${context.teamSiteDomain}` +
                    `&ImgPath=https://${context.teamSiteDomain}${template.ServerRelativeUrl}`;

                // console.log(eachTemplate.imgUrl);

            }
            if (template.Name.indexOf(".xls") > -1 || template.Name.indexOf(".xlts") > -1) {
                eachTemplate.altImgUrl = DEFAULT_XLSX_URL;
            }
            else if (template.Name.indexOf(".ppt") > -1 || template.Name.indexOf(".potx") > -1) {
                eachTemplate.altImgUrl = DEFAULT_PPTX_URL;

            }
            else if (template.Name.indexOf(".doc") > -1 || template.Name.indexOf(".dotx") > -1) {
                eachTemplate.altImgUrl = DEFAULT_DOCX_URL;
            }
        }

        currentTemplates.push(eachTemplate);
    });

    return currentTemplates.sort(compare);
}
