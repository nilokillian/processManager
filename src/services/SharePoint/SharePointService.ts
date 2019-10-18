import { WebPartContext } from "@microsoft/sp-webpart-base";
import { EnvironmentType } from "@microsoft/sp-core-library";
import {
  sp,
  SearchQuery,
  SearchResults,
  ItemAddResult,
  Web,
  AttachmentFileInfo,
  PermissionKind,
  ClientSideText,
  ClientSidePage
} from "@pnp/sp";
import {
  SPHttpClient,
  HttpClient,
  IHttpClientOptions
} from "@microsoft/sp-http";
import { IListCollection } from "./IList";
import { IListItemCollection } from "./IListItem";
import { IListFieldCollection } from "./IListField";
import { IListView } from "./IListView";
import { IChoiceFieldCollection } from "./IChoiceField";
import { SiteUser } from "@pnp/sp/src/siteusers";

export class SharePointServiceManager {
  public context: WebPartContext;
  public environmentType: EnvironmentType;

  public setup(context: WebPartContext): void {
    this.context = context;
  }

  public pnp_setup(content: WebPartContext): void {
    sp.setup({ spfxContext: content });
  }

  public pnp_getPolicy = (listTitle: string, itemId: number) => {
    const result = this.get_v2(
      `/_api/web/lists/getbytitle('${listTitle}')/items(${itemId})`
    );

    return result;
  };

  public pnp_getPolicies = (listTitle: string) => {
    const result = this.get_v2(
      `/_api/web/lists/getbytitle('${listTitle}')/items?`
    );

    return result;
  };

  public pnp_getListItems = (listTitle: string) => {
    const result = this.get_v2(
      `/_api/web/lists/getbytitle('${listTitle}')/items?`
    );

    return result;
  };

  public pnp_activeTasks = (
    listTitle: string,
    policyTitle: string,
    userGroupTitle?: string
  ) => {
    const result = this.get_v2(
      `/_api/web/lists/getbytitle('${listTitle}')/items?$filter=Policy eq '${policyTitle}'`
    );
    //UserGroupTitle eq '${userGroupTitle}' AND
    return result;
  };

  public getPolicyPage = async (
    libName: string,
    folderName: string,
    fileName: string
  ): Promise<any> => {
    const serverRelativeUrl = this.context.pageContext.web.serverRelativeUrl;

    return await sp.web
      .getFileByServerRelativeUrl(
        `${serverRelativeUrl}/${libName}/${folderName}/${fileName}`
      )
      .get()
      .then(r => r);
  };

  public getPolicyPages = async (
    libName: string,
    folderName: string
  ): Promise<any> => {
    const serverRelativeUrl = this.context.pageContext.web.serverRelativeUrl;

    return await sp.web
      .getFolderByServerRelativeUrl(
        `${serverRelativeUrl}/${libName}/${folderName}`
      )
      .files.get()
      .then(res => (res = res));
  };

  public deletePolicyPage = async (
    libName: string,
    folderName: string,
    fileName: string
  ): Promise<any> => {
    const serverRelativeUrl = this.context.pageContext.web.serverRelativeUrl;

    return await sp.web
      .getFileByServerRelativeUrl(
        `${serverRelativeUrl}/${libName}/${folderName}/${fileName}`
      )
      .delete()
      .then(r => r);
  };

  public pnp_CreatePage = async (pageTitle: string) => {
    const page = await sp.web.addClientSidePage(`${pageTitle}.aspx`);
    console.log("page", page);

    await this.pnp_AddControls(page);

    return page;
  };

  public pnp_AddControls = async (page: ClientSidePage) => {
    // this code adds a section, and then adds a control to that section. The control is added to the section's defaultColumn, and if there are no columns a single
    // column of factor 12 is created as a default. Here we add the ClientSideText part
    page.addSection().addControl(new ClientSideText("Here is some text!"));

    // here we add a section, add two columns, and add a text control to the second section so it will appear on the right of the page
    // add and get a reference to a new section
    const section = page.addSection();

    // add a column of factor 6
    section.addColumn(6);

    // add and get a reference to a new column of factor 6
    const column = section.addColumn(6);

    // add a text control to the second new column
    column.addControl(
      new ClientSideText(
        "Be sure to check out the developer guide at https://github.com/SharePoint/PnP-JS-Core/wiki/Developer-Guide"
      )
    );

    // we need to save our content changes
    await page.save();
  };

  public pnp_LoadPage = async () => {
    const page = await ClientSidePage.fromFile(
      sp.web.getFileByServerRelativeUrl(
        "/sites/appAdminCenter/process-manager/SitePages/Template-1.aspx"
      )
    );
    console.log("page", page);
  };

  public async get_v2(
    relativeEndpintUrl: string,
    userQueryParams?: object
  ): Promise<any> {
    try {
      const response = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}${relativeEndpintUrl}`,
        SPHttpClient.configurations.v1,
        userQueryParams && { body: JSON.stringify(userQueryParams) }
      );
      return !response.ok
        ? Promise.reject("Get Request Faild")
        : response.json();
    } catch (error) {
      //   throw new Error(error);
      return Promise.reject(error);
    }
  }

  public flowTrigger(
    relativeEndpintUrl: string,
    userQueryParams?: object
  ): Promise<any> {
    return this.context.spHttpClient
      .post(`${relativeEndpintUrl}`, SPHttpClient.configurations.v1, {
        body: JSON.stringify(userQueryParams)
      })
      .then(response => {
        return response.json();
      })
      .catch(error => {
        return Promise.reject(error);
      });
  }

  public post_flow(
    relativeEndpintUrl: string,
    httpClientOptions: IHttpClientOptions
  ) {
    this.context.httpClient
      .post(relativeEndpintUrl, HttpClient.configurations.v1, httpClientOptions)
      .then(res => {
        return res;
      });
  }

  public pnp_hasEditRights() {
    let web = new Web(this.context.pageContext.web.absoluteUrl);
    return web
      .currentUserHasPermissions(PermissionKind.EditListItems)
      .then(res => {
        return res;
      });
  }

  public pnp_hasDeleteRights() {
    let web = new Web(this.context.pageContext.web.absoluteUrl);

    return web
      .currentUserHasPermissions(PermissionKind.DeleteListItems)
      .then(res => {
        return res;
      });
  }

  public pnp_hasDeleteRightsV2(loginName) {
    let web = new Web(this.context.pageContext.web.absoluteUrl);
    return web
      .userHasPermissions(loginName, PermissionKind.DeleteListItems)
      .then(perms => {
        return perms;
      });
  }

  public pnp_hasEditRightsV2(loginName) {
    let web = new Web(this.context.pageContext.web.absoluteUrl);
    return web
      .userHasPermissions(loginName, PermissionKind.EditListItems)
      .then(perms => {
        return perms;
      });
  }

  public pnp_getItem(
    listTitle: string,
    itemId: number,
    fieldExpand?: string[],
    fieldSelect?: string[]
  ): Promise<any> {
    const web = new Web(this.context.pageContext.web.absoluteUrl);
    return web.lists
      .getByTitle(listTitle)
      .items.getById(itemId)
      .select(...fieldSelect)
      .expand(...fieldExpand)
      .get()
      .then((item: any) => {
        return item;
      })
      .catch(error => {
        return Promise.reject(error);
      });
  }

  public pnp_getItemsbyId(
    listId: string,
    fieldExpand?: string[],
    fieldSelect?: string[]
  ): Promise<any> {
    const web = new Web(this.context.pageContext.web.absoluteUrl);
    return web.lists
      .getById(listId)
      .items.select(...fieldSelect)
      .expand(...fieldExpand)
      .get()
      .then((items: any) => {
        return items;
      })
      .catch(error => {
        return Promise.reject(error);
      });
  }

  public pnp_getUserId = async (userEmail: string) => {
    const oUser = await sp.web.ensureUser(userEmail);

    return oUser.data;
  };

  public pnp_createList = async (listTitle: string, description?: string) => {
    const list = await sp.web.lists.add(listTitle, description, 100, true);
    // {
    //   Hidden: true
    // }
    return list;
  };

  public pnp_createField = async (listTitle: string, fields?: string[]) => {
    let isOk: boolean;
    const list = sp.web.lists.getByTitle(listTitle);
    const batch = sp.web.createBatch();

    fields.forEach(async field => {
      list.fields
        .inBatch(batch)
        .addText(field)
        .then(f => f);

      list.views
        .getByTitle("All Items")
        .fields.inBatch(batch)
        .add(field);
    });

    return await batch
      .execute()
      .then(() => {
        isOk = true;
      })
      .catch(error => {
        console.log(error);
      });
  };

  public pnp_getItemsByTitle(
    listTitle: string,
    fieldExpand?: string[],
    fieldSelect?: string[]
  ): Promise<any> {
    return sp.web.lists
      .getByTitle(listTitle)
      .items.select(...fieldSelect)
      .expand(...fieldExpand)
      .get()
      .then((items: any) => {
        return items;
      })
      .catch(error => {
        return Promise.reject(error);
      });
  }

  public pnp_createGroup() {
    return sp.web.siteGroups
      .add({
        Title: "trew345",
        Owner: { Id: 1073741823, PrincipalType: 1, Title: "System Account" }
      })
      .then(g => {
        console.log(g);
      });
  }

  public async pnp_createGroupV2(groupName: string, ownerLoginName?: string) {
    //const oUser = await sp.web.ensureUser(ownerLoginName);

    return sp.web.siteGroups
      .add({
        Title: groupName
        ///Owner: { id: oUser.data.Id }
      })
      .then(g => {
        console.log(g);
      });
  }
  public async pnp_deleteGroupMember(groupId: number, userId: number) {
    return await sp.web.siteGroups.getById(groupId).users.removeById(userId);
  }

  public async pnp_deleteGroup(groupId: number) {
    return await sp.web.siteGroups.removeById(groupId);
  }

  public async pnp_addGroupMember(groupId: number, userLoginNames: string[]) {
    let isOk: boolean;
    const batch = sp.web.createBatch();
    userLoginNames.map(userLogin => {
      sp.web.siteGroups
        .getById(groupId)
        .users.inBatch(batch)
        .add(userLogin)
        .then((value: SiteUser) => {
          console.log(value);
        });
    });

    await batch
      .execute()
      .then(() => {
        isOk = true;
      })
      .catch(error => {
        console.log(error);
      });
  }

  public async pnp_getGroupOwner() {
    sp.web.siteGroups
      .select(
        "Id",
        "Name",
        "Description",
        "Owner/Id",
        "Owner/Title",
        "Owner/PrincipalType"
      )
      .expand("Owner")
      .get()
      .then(console.log);
  }

  public pnp_getGroups() {
    return sp.web.siteGroups.get().then(res => res);
  }

  // public pnp_getGroupMembers(groups: any[]) {
  //   let users: any[] = [];

  //   groups.map(async group => {
  //     const currrentUsers = await sp.web.siteGroups
  //       .getById(Number(group.id))
  //       .users.get();

  //     users.push(
  //       ...currrentUsers.map(currentUser => {
  //         return {
  //           title: currentUser.Title,
  //           email: currentUser.Email,
  //           userId: currentUser.Id,
  //           groupName: group.name
  //         };
  //       })
  //     );
  //   });
  //   console.log("users", users);
  //   return users;
  //   // return sp.web.siteGroups.getById(groupId).users.        get();
  // }

  public async pnp_getGroupMembers(groups: any[]) {
    const users: any[] = [];

    for (let i = 0; i < groups.length; i++) {
      if (groups[i].id !== 0) {
        const currrentUsers = await sp.web.siteGroups
          .getById(Number(groups[i].id))
          .users.get();

        currrentUsers.map(currentUser => {
          users.push({
            title: currentUser.Title,
            email: currentUser.Email,
            userId: currentUser.Id,
            groupName: groups[i].name
          });
        });
      }
    }

    console.log("users", users);
    return users;
  }

  public pnp_post(listId: string, objectofValues?: object): Promise<any> {
    return sp.web.lists
      .getById(listId)
      .items.add(objectofValues)
      .then((response: ItemAddResult) => {
        return response.data;
      })
      .catch(error => {
        return Promise.reject(error);
      });
  }

  public pnp_postByTitle(
    listTitle: string,
    objectofValues?: object
  ): Promise<any> {
    return sp.web.lists
      .getByTitle(listTitle)
      .items.add(objectofValues)
      .then((response: ItemAddResult) => {
        return response.data;
      })
      .catch(error => {
        return Promise.reject(error);
      });
  }

  public pnp_post_multiple(listId: string, items: any[]): Promise<any> {
    const list = sp.web.lists.getById(listId);

    return list.getListItemEntityTypeFullName().then(entityTypeFullName => {
      let batch = sp.web.createBatch();

      items.map(item => {
        list.items
          .inBatch(batch)
          .add(item, entityTypeFullName)
          .then(b => {
            // console.log(b);
          });
      });

      return batch.execute().then(() => {
        return batch;
      });
    });
  }

  public pnp_postByTitle_multiple(
    listTitle: string,
    items: any[]
  ): Promise<any> {
    const list = sp.web.lists.getByTitle(listTitle);

    return list.getListItemEntityTypeFullName().then(entityTypeFullName => {
      let batch = sp.web.createBatch();

      items.map(item => {
        list.items
          .inBatch(batch)
          .add(item, entityTypeFullName)
          .then(b => {
            // console.log(b);
          });
      });

      return batch.execute().then(() => {
        return batch;
      });
    });
  }
  public pnp_update_collection_filter(
    listTitle: string,
    fieldFilterBy: string,
    filterValue: string,
    fieldToUpdate: string,
    newValue: string
  ) {
    // you are getting back a collection here
    sp.web.lists
      .getByTitle(listTitle)
      .items.filter(`${fieldFilterBy} eq '${filterValue}'`)
      .get()
      .then((items: any[]) => {
        // see if we got something
        if (items.length > 0) {
          sp.web.lists
            .getByTitle(listTitle)
            .items.getById(items[0].Id)
            .update({
              [fieldToUpdate]: newValue
            })
            .then(result => {
              // here you will have updated the item
              console.log(JSON.stringify(result));
            });
        }
      });
  }

  public pnp_update_multiple(
    listTitle: string,
    itemIds: any[],
    items: any[]
  ): Promise<any> {
    let web = new Web(this.context.pageContext.web.absoluteUrl);
    let list = web.lists.getByTitle(listTitle);

    return list.getListItemEntityTypeFullName().then(entityTypeFullName => {
      let batch = web.createBatch();

      items.map(item => {
        itemIds.map(id => {
          list.items
            .getById(id)
            .inBatch(batch)
            .update(item, "*", entityTypeFullName)
            .then(b => {
              // console.log(b);
            });
        });
      });
      // note requirement of "*" eTag param - or use a specific eTag value as needed

      batch.execute().then(d => console.log("Done"));
    });
  }

  public pnp_delete(listTitle: string, itemId: number): Promise<any> {
    return sp.web.lists
      .getByTitle(listTitle)
      .items.getById(itemId)
      .delete()
      .then(res => {
        return res;
      })
      .catch(err => {
        return Promise.reject(err);
      });
  }

  public pnp_delete_multiple(
    listTitle: string,
    itemIds: number[]
  ): Promise<any> {
    const list = sp.web.lists.getByTitle(listTitle);

    return list.getListItemEntityTypeFullName().then(entityTypeFullName => {
      let batch = sp.web.createBatch();

      itemIds.map(id => {
        list.items
          .getById(id)
          .inBatch(batch)
          .delete()
          .then(b => {
            // console.log(b);
          });
      });

      // note requirement of "*" eTag param - or use a specific eTag value as needed

      batch.execute().then(d => console.log("Done"));
    });

    // return sp.web.lists
    //   .getByTitle(listTitle)
    //   .items.getById(itemId)
    //   .inBatch(batch)
    //   .delete()
    //   .then(res => {
    //     return res;
    //   })
    //   .catch(err => {
    //     return Promise.reject(err);
    //   });
  }

  public pnp_update(
    listId: string,
    itemId: number,
    objectofValues?: object
  ): Promise<any> {
    return sp.web.lists
      .getById(listId)
      .items.getById(itemId)
      .update(objectofValues)
      .then((response: ItemAddResult) => {
        return response.item;
      })
      .catch(error => console.log(error));
  }

  public pnp_updateByTitle(
    listTitle: string,
    itemId: number,
    objectofValues?: object
  ): Promise<any> {
    return sp.web.lists
      .getByTitle(listTitle)
      .items.getById(itemId)
      .update(objectofValues)
      .then((response: ItemAddResult) => {
        return response.item;
      })
      .catch(error => console.log(error));
  }

  public pnp_addAttachment(
    listId: string,
    itemId: number,
    attachments
  ): Promise<any> {
    let web = new Web(this.context.pageContext.web.absoluteUrl);
    const filename = attachments.fileName;
    const data = attachments.data;
    const item = web.lists.getById(listId).items.getById(itemId);
    return item.attachmentFiles.add(filename, data).then(v => {
      //console.log(v);
      return v;
    });
  }

  public pnp_addAttachment_multiple(
    listId: string,
    itemId: number,
    attachments
  ): Promise<any> {
    const web = new Web(this.context.pageContext.web.absoluteUrl);

    const list = web.lists.getById(listId);
    const fileInfos: AttachmentFileInfo[] = [];

    attachments.map(file => {
      fileInfos.push({
        name: file.fileName,
        content: file.data
      });
    });

    return list.items
      .getById(itemId)
      .attachmentFiles.addMultiple(fileInfos)
      .then(res => {
        console.log(res);
        return res;
      });
  }

  public pnp_getAttachment(listId: string, itemId: number): Promise<any> {
    const web = new Web(this.context.pageContext.web.absoluteUrl);
    const item = web.lists.getById(listId).items.getById(itemId);

    // get all the attachments
    return item.attachmentFiles.get().then(v => {
      return v;
    });
  }

  public pnp_deleteAttachment(
    listId: string,
    itemId: number,
    attachments
  ): Promise<any> {
    let web = new Web(this.context.pageContext.web.absoluteUrl);
    const item = web.lists.getById(listId).items.getById(itemId);

    return item.attachmentFiles
      .getByName(attachments.fileName)
      .delete()
      .then(v => {
        return v;
      });
  }

  public pnp_deleteAttachments(
    listId: string,
    itemId: number,
    attachments
  ): Promise<any> {
    const web = new Web(this.context.pageContext.web.absoluteUrl);
    const list = web.lists.getById(listId);
    // const item = web.lists.getById(listId).items.getById(itemId);

    return list.items
      .getById(itemId)
      .attachmentFiles.deleteMultiple(...attachments)
      .then(r => {
        console.log(r);
        return r;
      });
  }

  public pnp_unpdateAttachment(): Promise<any> {
    let web = new Web(this.context.pageContext.web.absoluteUrl);
    let item = web.lists.getByTitle("MyList").items.getById(1);

    return item.attachmentFiles
      .getByName("file2.txt")
      .setContent("My new content!!!")
      .then(v => {
        console.log(v);
      });
  }

  public async getBinary(serverRelativeUrl: string): Promise<any> {
    // try {
    //   const response = await this.context.spHttpClient.get(
    //     `${
    //       this.context.pageContext.web.absoluteUrl
    //     }/_api/web/getfilebyserverrelativeurl('${serverRelativeUrl}')/$value?binaryStringResponseBody=true`,
    //     SPHttpClient.configurations.v1
    //   );
    //   return !response.ok
    //     ? Promise.reject("Get Request Faild")
    //     : response.body
    //         //response.body  => issue with body in 1.8.2 does not exist
    //         .getReader()
    //         .read()
    //         .then(r => r);
    // } catch (error) {
    //   //   throw new Error(error);
    //   return Promise.reject(error);
    // }
  }

  public getFieldsFromView(
    listId: string,
    selectedViewId?: string
  ): Promise<IListView> {
    if (!listId) return;

    return this.get_v2(
      `/_api/web/lists/getbyid('${listId}')/views?${
        !selectedViewId
          ? "$filter=DefaultView eq true"
          : "$filter=Id eq selectedViewId"
      }`
    );
  }

  public getItemsBySharePointAPI(
    listId: string,
    selectedFields?: string[],
    expend?: string[],
    filters?: string
  ): Promise<IListItemCollection> {
    return this.get_v2(
      `/_api/web/lists/getbyid('${listId}')/items?${
        selectedFields.length !== 0 ? `$select=${selectedFields.join(",")}` : ""
      } ${expend ? `&$expand=${expend.join(",")}` : ""}${
        filters ? `&$filter=${filters}` : ""
      }&$top=5000`
    );
  }
  //Project ID: 394
  public async getItemsBySearchAPI(query?: string) {
    const q: SearchQuery = {
      Querytext: `Path:${this.context.pageContext.web.absoluteUrl}*`,
      RowLimit: 100,
      EnableInterleaving: true
    };
    let y: SearchQuery;

    // define a search query object matching the SearchQuery interface
    const result: SearchResults = await sp.search(q);
    console.log("result :", result);
  }

  public getListFields(
    listId: string,
    showHiddenFields: boolean = false
  ): Promise<IListFieldCollection> {
    return this.get_v2(
      `/_api/web/lists/getById('${listId}')/fields${
        !showHiddenFields ? "?$filter=Hidden eq false" : ""
        // ? "?$filter=Hidden eq false and ReadOnlyField eq false"
      }`
    );
  }

  public getChoiseOptions(listId: string, fieldName: string): Promise<any> {
    const url = `/_api/web/lists/getById('${listId}')/fields?$filter=EntityPropertyName eq '${fieldName}'`;

    return this.get_v2(url)
      .then((res: IChoiceFieldCollection) => {
        return res.value;
      })
      .catch(error => console.log(error));
  }

  public createExpendedFields = (fieldOptions: any[]): string[] => {
    const expendedFields = [];
    for (let field in fieldOptions) {
      if (
        fieldOptions[field].fieldType === "User" ||
        fieldOptions[field].fieldType === "UserMulti" ||
        fieldOptions[field].fieldType === "Lookup" ||
        fieldOptions[field].fieldType === "LookupMulti" ||
        fieldOptions[field].fieldType === "Attachments"
      ) {
        expendedFields.push(fieldOptions[field].key);
      }
    }

    return expendedFields;
  };

  public createQueriedFields = (fieldOptions: any[]): string[] => {
    const queriedFields = [];

    for (let field in fieldOptions) {
      switch (fieldOptions[field].fieldType) {
        case "User": {
          queriedFields.push(
            fieldOptions[field].key + "/Title",
            fieldOptions[field].key + "/EMail",
            fieldOptions[field].key + "/ID"
          );
          break;
        }
        case "Lookup": {
          queriedFields.push(
            fieldOptions[field].key + "/" + fieldOptions[field]["lookupField"]
          );
          break;
        }
        case "LookupMulti": {
          queriedFields.push(
            fieldOptions[field].key + "/" + fieldOptions[field]["lookupField"]
          );
          break;
        }

        case "Attachments": {
          queriedFields.push("AttachmentFiles");
          break;
        }

        case "UserMulti": {
          queriedFields.push(
            fieldOptions[field].key + "/Title",
            // fieldOptions[field].key + "/EMail",
            fieldOptions[field].key + "/ID"
          );

          break;
        }

        default: {
          queriedFields.push(fieldOptions[field].key);

          break;
        }
      }
    }

    return queriedFields;
  };
}

const SharePointService = new SharePointServiceManager();

export default SharePointService;
