/* eslint-disable no-unused-expressions */
/* eslint-disable no-var */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable react/jsx-no-target-blank */
/* eslint-disable @typescript-eslint/no-empty-function */
/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable @typescript-eslint/no-use-before-define */
/* eslint-disable eqeqeq */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @rushstack/no-new-null */
/* eslint-disable prefer-const */
/* eslint-disable @typescript-eslint/no-var-requires */
/* eslint-disable @typescript-eslint/no-explicit-any */
import { sp } from "@pnp/sp/presets/all";
import {
  IFilter,
  IListItems,
  IListItemUsingId,
  IAddList,
  IUpdateList,
  ISPList,
  IDetailsListGroup,
  ISPAttachment,
  IAttachDelete,
  ISPListChoiceField,
  ISPDocument,
  ISPAddDocument,
  ISPUpdateDocument,
  IDeleteDocument,
  ISPGrpMember,
  IAnotherListItems,
  IAnotherAddList,
  IAnotherUpdateList,
  IAnotherDeleteList,
  IAnotherListItemUsingId,
  IGetDocLibFiles,
  IDocFiles,
  IAddDocLibFiles,
  IInsertFiles,
  IRemoveFiles,
  IAzureUsers,
} from "./ISPServicesProps";
import { IItemAddResult, IItemUpdateResult } from "@pnp/sp/items";
import "@pnp/sp/webs";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups/web";
import { MSGraphClient } from "@microsoft/sp-http";

const getAllUsers = async (): Promise<[]> => {
  return await sp.web.siteUsers();
};

const SPAddItem = async (params: IAddList): Promise<IItemAddResult> => {
  return await sp.web.lists
    .getByTitle(params.Listname)
    .items.add(params.RequestJSON);
};

const SPUpdateItem = async (params: IUpdateList): Promise<IItemAddResult> => {
  return await sp.web.lists
    .getByTitle(params.Listname)
    .items.getById(params.ID)
    .update(params.RequestJSON);
};

const SPDeleteItem = async (params: ISPList): Promise<string> => {
  return await sp.web.lists
    .getByTitle(params.Listname)
    .items.getById(params.ID)
    .recycle();
};

const SPReadItems = async (params: IListItems): Promise<[]> => {
  params = formatInputs(params);
  let filterValue: string = formatFilterValue(
    params.Filter || [],
    params.FilterCondition ? params.FilterCondition : "",
  );

  return await sp.web.lists
    .getByTitle(params.Listname)
    .items.select(params.Select || "*")
    .filter(filterValue)
    .expand(params.Expand || "")
    .top(params.Topcount || 0)
    .orderBy(params.Orderby || "ID", params.Orderbydecorasc)
    .get();
};

const SPReadItemUsingId = async (params: IListItemUsingId): Promise<[]> => {
  return await sp.web.lists
    .getByTitle(params.Listname)
    .items.getById(params.SelectedId)
    .select(params.Select ? params.Select : "")
    .expand(params.Expand ? params.Expand : "")
    .get();
};

const SPAddAttachments = async (params: ISPAttachment) => {
  const files: any[] = params.Attachments;
  return await sp.web.lists
    .getByTitle(params.ListName)
    .items.getById(params.ListID)
    .attachmentFiles.addMultiple(files);
};

const SPGetAttachments = async (params: ISPList) => {
  const item: any = sp.web.lists
    .getByTitle(params.Listname)
    .items.getById(params.ID);
  return await item.attachmentFiles();
};

const SPDeleteAttachments = async (params: IAttachDelete) => {
  return await sp.web.lists
    .getByTitle(params.ListName)
    .items.getById(params.ListID)
    .attachmentFiles.getByName(params.AttachmentName)
    .delete();
};

const SPGetChoices = async (params: ISPListChoiceField) => {
  return await sp.web.lists
    .getByTitle(params.Listname)
    .fields.getByInternalNameOrTitle(params.FieldName)
    .get();
};

const SPDetailsListGroupItems = async (params: IDetailsListGroup) => {
  type RecordType = { [key: string]: any; indexValue: number };
  let newRecords: RecordType[] = [];
  params.Data.forEach((arr, index) => {
    return newRecords.push({
      Lesson: arr[params.Column],
      indexValue: index,
    });
  });

  let varGroup: { key: any; name: any; startIndex: number; count: number }[] =
    [];
  let UniqueRecords = newRecords.reduce(function (item, e1) {
    var matches = item.filter(function (e2) {
      return e1.Lesson === e2.Lesson;
    });

    if (matches.length == 0) {
      item.push(e1);
    }
    return item;
  }, [] as RecordType[]);

  UniqueRecords.forEach((ur) => {
    let recordLength = newRecords.filter((arr) => {
      return arr.Lesson == ur.Lesson;
    }).length;
    varGroup.push({
      key: ur.Lesson,
      name: ur.Lesson,
      startIndex: ur.indexValue,
      count: recordLength,
    });
  });
  return varGroup;
};

const batchInsert = async (params: {
  ListName: string;
  responseData: any[];
}): Promise<any> => {
  const list = sp.web.lists.getByTitle(params.ListName);
  const batch = sp.web.createBatch();
  const promises: any[] = [];

  for (const data of params.responseData) {
    const promise = list.items.inBatch(batch).add(data);
    promises.push(promise);
  }

  await batch
    .execute()
    .then(() => {
      return promises;
    })
    .catch((error) => console.log(error));
};

const batchUpdate = async (params: {
  ListName: string;
  responseData: any[];
}): Promise<any> => {
  const list = sp.web.lists.getByTitle(params.ListName);
  const batch = sp.web.createBatch();
  const promises: any[] = [];

  for (const data of params.responseData) {
    const promise: any = list.items
      .getById(data.ID)
      .inBatch(batch)
      .update(data);
    promises.push(promise);
  }

  await batch
    .execute()
    .then(() => {
      return promises;
    })
    .catch((error) => console.log(error));
};

const batchDelete = async (params: {
  ListName: string;
  responseData: any[];
}): Promise<any> => {
  const list = sp.web.lists.getByTitle(params.ListName);
  const batch = sp.web.createBatch();
  const promises: Promise<IItemUpdateResult>[] = [];

  for (const data of params.responseData) {
    const promise: any = list.items.getById(data.ID).inBatch(batch).recycle();
    promises.push(promise);
  }

  await batch
    .execute()
    .then(() => {
      return promises;
    })
    .catch((error) => console.log(error));
};

const formatInputs = (data: IListItems): IListItems => {
  !data.Select ? (data.Select = "*") : "";
  !data.Topcount ? (data.Topcount = 5000) : "";
  !data.Orderby ? (data.Orderby = "ID") : "";
  !data.Expand ? (data.Expand = "") : "";
  data.Orderbydecorasc !== true && data.Orderbydecorasc !== false
    ? (data.Orderbydecorasc = true)
    : (data.Orderbydecorasc = false);
  !data.PageCount ? (data.PageCount = 10) : "";
  !data.PageNumber ? (data.PageNumber = 1) : "";

  return data;
};

const formatFilterValue = (
  params: IFilter[],
  filterCondition: string,
): string => {
  let strFilter: string = "";
  if (params) {
    for (let i = 0; i < params.length; i++) {
      if (params[i].FilterKey) {
        if (i != 0) {
          if (filterCondition == "and" || filterCondition == "or") {
            strFilter += " " + filterCondition + " ";
          } else {
            strFilter += " and ";
          }
        }

        if (
          params[i].Operator.toLocaleLowerCase() == "eq" ||
          params[i].Operator.toLocaleLowerCase() == "ne" ||
          params[i].Operator.toLocaleLowerCase() == "gt" ||
          params[i].Operator.toLocaleLowerCase() == "lt" ||
          params[i].Operator.toLocaleLowerCase() == "ge" ||
          params[i].Operator.toLocaleLowerCase() == "le"
        )
          strFilter +=
            params[i].FilterKey +
            " " +
            params[i].Operator +
            "'" +
            params[i].FilterValue +
            "'";
        else if (params[i].Operator.toLocaleLowerCase() == "substringof")
          strFilter +=
            params[i].Operator +
            "('" +
            params[i].FilterValue +
            "','" +
            params[i].FilterKey +
            "')";
      }
    }
  }
  return strFilter;
};

const SPReadDocumentItems = async (params: ISPDocument) => {
  const tempData: any[] = [];

  const folders: any = await sp.web
    .getFolderByServerRelativePath(params.DocumentName)
    .folders.get();
  tempData.push({ Folders: folders });

  const files = await sp.web
    .getFolderByServerRelativePath(params.DocumentName)
    .files.select(params.Select)
    .expand(params.Expand)
    .get();
  tempData.push({ AllData: files });

  return tempData;
};

const SPAddDocumentItem = async (params: ISPAddDocument) => {
  await sp.web
    .getFolderByServerRelativeUrl(params.DocumentName)
    .files.add(params.FileName, params.FileContent)
    .then(({ file }) => {
      file.getItem().then((item: any) => {
        item.update(params.ColumnNames);
      });
    });
};

const SPUpdateDocumentItem = async (params: ISPUpdateDocument) => {
  const file: any = sp.web.getFileByServerRelativeUrl(params.ServerUrl);
  return await sp.web
    .getFileByServerRelativeUrl(params.ServerUrl)()
    .then(() =>
      file
        .setContent(params.FileContent)
        .then((items: any) =>
          items.getItem().then((item: any) => item.update(params.UpdateData)),
        ),
    );
};

const SPDeleteDocumentItem = async (params: IDeleteDocument) => {
  sp.web.lists
    .getByTitle(params.DocumentName)
    .items.getById(params.ID)
    .delete();
};

const getSPGroupMember = async (params: ISPGrpMember): Promise<[]> => {
  return await sp.web.siteGroups.getByName(params.GroupName).users.get();
};

const AnotherformatInputs = (data: IAnotherListItems): IAnotherListItems => {
  !data.Select ? (data.Select = "*") : "";
  !data.Topcount ? (data.Topcount = 5000) : "";
  !data.Orderby ? (data.Orderby = "ID") : "";
  !data.Expand ? (data.Expand = "") : "";
  data.Orderbydecorasc !== true && data.Orderbydecorasc !== false
    ? (data.Orderbydecorasc = true)
    : "";
  !data.PageCount ? (data.PageCount = 10) : "";
  !data.PageNumber ? (data.PageNumber = 1) : "";
  return data;
};

const getAnotherSPReadItems = async (
  params: IAnotherListItems,
): Promise<[]> => {
  const WebUrl = Web(params.SiteUrl);

  params = AnotherformatInputs(params);
  let filterValue: string = formatFilterValue(
    params.Filter || [],
    params.FilterCondition ? params.FilterCondition : "",
  );
  return await WebUrl.lists
    .getByTitle(params.Listname)
    .items.select(params.Select || "*")
    .filter(filterValue)
    .expand(params.Expand || "")
    .top(params.Topcount || 0)
    .orderBy(params.Orderby || "ID", params.Orderbydecorasc)
    .get();
};

const AnotherSPAddItem = async (
  params: IAnotherAddList,
): Promise<IItemAddResult> => {
  const WebUrl = Web(params.SiteUrl);

  return await WebUrl.lists
    .getByTitle(params.Listname)
    .items.add(params.RequestJSON);
};

const AnotherSPUpdateItem = async (
  params: IAnotherUpdateList,
): Promise<IItemAddResult> => {
  const WebUrl = Web(params.SiteUrl);

  return await WebUrl.lists
    .getByTitle(params.Listname)
    .items.getById(params.ID)
    .update(params.RequestJSON);
};

const AnotherSPDeleteItem = async (
  params: IAnotherDeleteList,
): Promise<void> => {
  const WebUrl = Web(params.SiteUrl);

  return await WebUrl.lists
    .getByTitle(params.Listname)
    .items.getById(params.ID)
    .delete();
};

const AnotherSPReadItemUsingId = async (
  params: IAnotherListItemUsingId,
): Promise<object> => {
  const WebUrl = Web(params.SiteUrl);
  !params.Select ? (params.Select = "") : params.Select;
  !params.Expand ? (params.Expand = "") : params.Expand;
  return await WebUrl.lists
    .getByTitle(params.Listname)
    .items.getById(params.SelectedId)
    .select(params.Select)
    .expand(params.Expand)
    .get();
};

const getDocLibFiles = async (params: IGetDocLibFiles): Promise<object[]> => {
  let FilesArr: IDocFiles[] = [];
  await sp.web
    .getFolderByServerRelativePath(params.FilePath)
    .files.get()
    .then((DocRes) => {
      if (DocRes.length) {
        DocRes.forEach((item) => {
          FilesArr.push({
            name: item.Name,
            content: item,
            type: "Inlist",
          });
        });
      }
    })
    .catch((err) => console.log(err, "Get Document Library Files"));
  return FilesArr;
};

const addDocLibFiles = async (params: IAddDocLibFiles) => {
  let getFilePath: string = params.FilePath;
  if (params.Datas.length) {
    let delAttachments: IDocFiles[] = params.Datas.filter(
      (files: IDocFiles) => {
        return files.type == "Delete";
      },
    );

    let addAttachments: IDocFiles[] = params.Datas.filter(
      (files: IDocFiles) => {
        return files.type == "New";
      },
    );

    if (params.FolderNames.length) {
      for (let j: number = 0; j < params.FolderNames.length; j++) {
        await sp.web
          .getFolderByServerRelativePath(getFilePath)
          .folders.addUsingPath(params.FolderNames[j], true)
          .then(async (res) => {
            getFilePath = res.data.ServerRelativeUrl;
            if (j === params.FolderNames.length - 1 && delAttachments.length) {
              for (let k: number = 0; k < delAttachments.length; k++) {
                await sp.web
                  .getFolderByServerRelativePath(getFilePath)
                  .files.getByName(delAttachments[k].name)
                  .delete()
                  .then(async () => {
                    if (
                      addAttachments.length &&
                      delAttachments.length - 1 === k
                    ) {
                      for (let i: number = 0; i < addAttachments.length; i++) {
                        await sp.web
                          .getFolderByServerRelativePath(getFilePath)
                          .files.addUsingPath(
                            addAttachments[i].name,
                            addAttachments[i].content,
                            {
                              Overwrite: true,
                            },
                          )
                          .then(() => {})
                          .catch((error) => {
                            console.log("Error creating file", error);
                          });
                      }
                    }
                  })
                  .catch((error) => {
                    console.log("Delete  attachements", error);
                  });
              }
            } else if (
              j === params.FolderNames.length - 1 &&
              addAttachments.length
            ) {
              for (let z: number = 0; z < addAttachments.length; z++) {
                await sp.web
                  .getFolderByServerRelativePath(getFilePath)
                  .files.addUsingPath(
                    addAttachments[z].name,
                    addAttachments[z].content,
                    {
                      Overwrite: true,
                    },
                  )
                  .then(() => {})
                  .catch((error) => {
                    console.log("Error creating file", error);
                  });
              }
            }
          })
          .catch((err) => console.log("creating folder", err));
      }
    } else {
      getFilePath = params.FilePath;
      if (delAttachments.length) {
        for (let i: number = 0; i < delAttachments.length; i++) {
          await sp.web
            .getFolderByServerRelativePath(getFilePath)
            .files.getByName(delAttachments[i].name)
            .delete()
            .then(async () => {
              if (addAttachments.length && delAttachments.length - 1 === i) {
                for (let j: number = 0; j < addAttachments.length; j++) {
                  await sp.web
                    .getFolderByServerRelativePath(getFilePath)
                    .files.addUsingPath(
                      addAttachments[j].name,
                      addAttachments[j].content,
                      {
                        Overwrite: true,
                      },
                    )
                    .then(() => {})
                    .catch((error) => {
                      console.log("Error creating file", error);
                    });
                }
              }
            })
            .catch((error) => console.log("Delete attachements", error));
        }
      } else if (addAttachments.length) {
        for (let i: number = 0; i < addAttachments.length; i++) {
          await sp.web
            .getFolderByServerRelativePath(params.FilePath)
            .files.addUsingPath(
              addAttachments[i].name,
              addAttachments[i].content,
              {
                Overwrite: true,
              },
            )
            .then(() => {})
            .catch((error) => {
              console.log("Error creating file", error);
            });
        }
      }
    }
  }
  return getFilePath ? getDocLibFiles({ FilePath: getFilePath }) : [];
};

const fileInsert = (params: IInsertFiles) => {
  if (params.files.length) {
    let insertFile: IDocFiles[] = [...params.data];
    for (let i = 0; i < params.files.length; i++) {
      if (
        ![...insertFile].some((values) => {
          return (
            values.name == params.files[i].name && values.type !== "Delete"
          );
        })
      )
        insertFile.push({
          name: params.files[i].name,
          content: params.files[i],
          type: "New",
        });
    }
    return insertFile;
  }
};

const fileRemove = (params: IRemoveFiles) => {
  let removefiles: IDocFiles[] = [...params.data];
  let item = removefiles[params.index];
  if (item.type == "Inlist") {
    removefiles[params.index] = { ...item, ["type"]: "Delete" };
  } else if (item.type == "New") {
    removefiles.splice(params.index, 1);
  }
  return removefiles;
};

const GetAzureUsers = async (params: IAzureUsers): Promise<any[]> => {
  return await params.Context._msGraphClientFactory
    .getClient()
    .then(async (UsersApi: MSGraphClient) => {
      return await UsersApi.api("Users").top(999).get();
    });
};

const GetAzureUsersGroups = async (params: IAzureUsers): Promise<any[]> => {
  return await params.Context._msGraphClientFactory
    .getClient()
    .then(async (UsersApi: MSGraphClient) => {
      return await UsersApi.api("/groups").top(999).get();
    });
};

const SPReadItemVersionHistory = async (
  params: IListItemUsingId,
): Promise<[]> => {
  return await sp.web.lists
    .getByTitle(params.Listname)
    .items.getById(params.SelectedId)
    .versions.select(params.Select || "*")
    .expand(params.Expand || "")
    .get();
};

export default {
  getAllUsers,
  SPAddItem,
  SPUpdateItem,
  SPDeleteItem,
  SPReadItems,
  SPDetailsListGroupItems,
  SPGetChoices,
  SPAddAttachments,
  SPGetAttachments,
  SPDeleteAttachments,
  SPReadItemUsingId,
  batchInsert,
  batchUpdate,
  batchDelete,
  SPReadDocumentItems,
  SPAddDocumentItem,
  SPUpdateDocumentItem,
  SPDeleteDocumentItem,
  getSPGroupMember,
  getAnotherSPReadItems,
  AnotherSPAddItem,
  AnotherSPUpdateItem,
  AnotherSPDeleteItem,
  AnotherSPReadItemUsingId,
  getDocLibFiles,
  addDocLibFiles,
  fileInsert,
  fileRemove,
  GetAzureUsers,
  GetAzureUsersGroups,
  SPReadItemVersionHistory,
};
