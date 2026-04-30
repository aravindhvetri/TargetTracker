/* eslint-disable @typescript-eslint/no-var-requires */
/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable eqeqeq */
/* eslint-disable @typescript-eslint/no-use-before-define */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable prefer-const */
/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from "react";
import { useState, useEffect } from "react";
import SPServices from "../../../../External/CommonServices/SPServices";
import { Config } from "../../../../External/CommonServices/Config";
import { DataTable } from "primereact/datatable";
import { Column } from "primereact/column";
import {
  multiPeoplePickerTemplate,
  peoplePickerTemplate,
} from "../../../../External/CommonTemplates/CommonTemplates";
import styles from "./TargetConfiguration.module.scss";
import { Button } from "primereact/button";
import { Dialog } from "primereact/dialog";
import { ConfirmDialog, confirmDialog } from "primereact/confirmdialog";
import { InputText } from "primereact/inputtext";
import { Dropdown } from "primereact/dropdown";
import {
  IBasePickerSuggestionsProps,
  IPersonaProps,
  NormalPeoplePicker,
} from "@fluentui/react";

type IFormErrors = {
  Technology: boolean;
  Target: boolean;
  ProjectManager: boolean;
  DeliveryHead: boolean;
};

const TargetConfig = (props: any) => {
  //States:
  const [targetConfigData, setTargetConfigData] = useState<
    ITargetConfigurationData[]
  >([]);
  const [showDialog, setShowDialog] = useState(false);
  const [isEdit, setIsEdit] = useState(false);
  const [formData, setFormData] = useState<ITargetConfigurationData>({
    ...Config.TargetConfigurationData,
  });
  const [technologyOptions, setTechnologyOptions] = useState<
    { name: string }[]
  >([]);
  const [siteUsers, setSiteUsers] = useState<IPersonaProps[]>([]);
  const [formErrors, setFormErrors] = useState<IFormErrors>({
    Technology: false,
    Target: false,
    ProjectManager: false,
    DeliveryHead: false,
  });

  //Use Effect to get Target Configuration Data:
  useEffect(() => {
    getTargetConfigurationData();
    getSiteUsers();
  }, []);

  //Get Target Configuration Data function:
  const getTargetConfigurationData = async () => {
    SPServices.SPReadItems({
      Listname: Config.ListNames.TargetConfiguration,
      Select:
        "*,ProjectManager/Id,ProjectManager/EMail,ProjectManager/Title,DeliveryHead/Id,DeliveryHead/EMail,DeliveryHead/Title",
      Expand: "ProjectManager,DeliveryHead",
      Orderby: "ID",
      Orderbydecorasc: true,
    })
      .then((response: any) => {
        let data: ITargetConfigurationData[] = [];
        response.forEach((items: any) => {
          if (items?.IsDelete) return;

          let _ProjectManager: IPeoplePickerDetails[] = [];
          if (items?.ProjectManager) {
            items?.ProjectManager.forEach((user: any) => {
              _ProjectManager.push({
                id: user?.Id,
                name: user?.Title,
                email: user?.EMail,
              });
            });
          }
          let _DeliveryHead: IPeoplePickerDetails[] = [];
          if (items?.DeliveryHead) {
            items?.DeliveryHead.forEach((user: any) => {
              _DeliveryHead.push({
                id: user?.Id,
                name: user?.Title,
                email: user?.EMail,
              });
            });
          }
          data.push({
            ID: items.ID || 0,
            Technology: items.Technology || "",
            Target: items.Target || "",
            ProjectManager: _ProjectManager || [],
            DeliveryHead: _DeliveryHead || [],
            IsDelete: items.IsDelete || false,
          });
        });
        setTargetConfigData([...data]);
        getTechnologyChoices();
      })
      .catch((error: any) => {
        console.log(error);
      });
  };

  const getSiteUsers = async () => {
    try {
      const users: any[] = await SPServices.getAllUsers();
      const mappedUsers: IPersonaProps[] = users
        .filter((user: any) => !!user?.Email)
        .map((user: any) => ({
          key: String(user?.Id || user?.LoginName || user?.Email),
          text: user?.Title || user?.Email || "",
          secondaryText: user?.Email || "",
        }));

      setSiteUsers(mappedUsers);
    } catch (error) {
      console.log(error);
      setSiteUsers([]);
    }
  };

  //Get Technology Choices function:
  const getTechnologyChoices = async () => {
    SPServices.SPGetChoices({
      Listname: Config.ListNames.TargetConfiguration,
      FieldName: "Technology",
    })
      .then((res: any) => {
        let temp: any[] = [];

        if (res?.Choices?.length) {
          res.Choices.forEach((val: any) => {
            temp.push({ name: val });
          });
        }
        setTechnologyOptions(temp);
      })
      .catch((error: any) => {
        console.log(error);
      });
  };

  //Render Manager Column function:
  const renderProjectManagerColumn = (rowData: ITargetConfigurationData) => {
    const projectManagers: IPeoplePickerDetails[] = rowData?.ProjectManager;
    return (
      <div>
        {rowData?.ProjectManager?.length > 1
          ? multiPeoplePickerTemplate(projectManagers)
          : peoplePickerTemplate(projectManagers[0])}
      </div>
    );
  };

  //Render Delivery Head Column function:
  const renderDeliveryHeadColumn = (rowData: ITargetConfigurationData) => {
    const deliveryHead: IPeoplePickerDetails[] = rowData?.DeliveryHead;
    return (
      <div>
        {rowData?.DeliveryHead?.length > 1
          ? multiPeoplePickerTemplate(deliveryHead)
          : peoplePickerTemplate(deliveryHead[0])}
      </div>
    );
  };

  //Render Action Column function:
  const renderActionColumn = (rowData: ITargetConfigurationData) => {
    return (
      <div className={styles.actionContainer}>
        <Button
          type="button"
          icon="pi pi-pencil"
          rounded
          text
          aria-label={`Edit ${rowData.Technology}`}
          className={`${styles.actionButton} ${styles.editButton}`}
          onClick={() => handleEdit(rowData)}
        />
        <Button
          type="button"
          icon="pi pi-trash"
          rounded
          text
          aria-label={`Delete ${rowData.Technology}`}
          className={`${styles.actionButton} ${styles.deleteButton}`}
          onClick={() => handleDelete(rowData)}
        />
      </div>
    );
  };

  //Handle Add Target function:
  const handleAdd = () => {
    setFormData({
      ...Config.TargetConfigurationData,
    });
    setFormErrors({
      Technology: false,
      Target: false,
      ProjectManager: false,
      DeliveryHead: false,
    });
    setIsEdit(false);
    setShowDialog(true);
  };

  const handleEdit = (rowData: ITargetConfigurationData) => {
    setFormData({ ...rowData });
    setFormErrors({
      Technology: false,
      Target: false,
      ProjectManager: false,
      DeliveryHead: false,
    });
    setIsEdit(true);
    setShowDialog(true);
  };

  //Handle Delete Target function:
  const handleDelete = async (rowData: ITargetConfigurationData) => {
    confirmDialog({
      header: "Delete target",
      message: `Are you sure you want to delete the ${rowData.Technology || "selected"} target configuration?`,
      icon: "pi pi-exclamation-triangle",
      acceptLabel: "Yes",
      rejectLabel: "No",
      acceptClassName: "p-button-danger",
      rejectClassName: "p-button-text",
      className: styles.deleteConfirmDialog,
      accept: async () => {
        await SPServices.SPUpdateItem({
          Listname: Config.ListNames.TargetConfiguration,
          ID: rowData.ID || 0,
          RequestJSON: {
            IsDelete: true,
          },
        });

        getTargetConfigurationData();
      },
    });
  };

  //Handle Save Target function:
  const handleSave = async () => {
    const validationErrors: IFormErrors = {
      Technology: !formData.Technology?.trim(),
      Target: !formData.Target?.trim(),
      ProjectManager: !formData.ProjectManager?.length,
      DeliveryHead: !formData.DeliveryHead?.length,
    };

    setFormErrors(validationErrors);

    if (Object.values(validationErrors).some((hasError) => hasError)) {
      return;
    }

    const json: any = {
      Technology: formData.Technology.trim(),
      Target: formData.Target.trim(),
      ProjectManagerId: {
        results: formData.ProjectManager.map((x) => x.id),
      },
      DeliveryHeadId: {
        results: formData.DeliveryHead.map((x) => x.id),
      },
      IsDelete: false,
    };

    if (isEdit) {
      await SPServices.SPUpdateItem({
        Listname: Config.ListNames.TargetConfiguration,
        ID: formData.ID || 0,
        RequestJSON: json,
      });
    } else {
      await SPServices.SPAddItem({
        Listname: Config.ListNames.TargetConfiguration,
        RequestJSON: json,
      });
    }

    setShowDialog(false);
    getTargetConfigurationData();
  };

  const suggestionsProps: IBasePickerSuggestionsProps = {
    suggestionsHeaderText: "Suggested users",
    noResultsFoundText: "No matching users",
  };

  const filterSiteUsers = (
    filterText: string,
    currentPersonas?: IPersonaProps[],
  ): IPersonaProps[] => {
    if (!filterText) return [];

    const selectedKeys = new Set(
      (currentPersonas || []).map((persona) => String(persona.key || "")),
    );
    const normalizedFilter = filterText.toLowerCase();

    return siteUsers.filter((persona) => {
      const personaKey = String(persona.key || "");
      const personaName = (persona.text || "").toLowerCase();
      const personaEmail = (persona.secondaryText || "").toLowerCase();

      return (
        !selectedKeys.has(personaKey) &&
        (personaName.indexOf(normalizedFilter) >= 0 ||
          personaEmail.indexOf(normalizedFilter) >= 0)
      );
    });
  };

  const getTextFromItem = (persona: IPersonaProps): string =>
    persona.text || "";

  const mapToPersona = (users: IPeoplePickerDetails[]): IPersonaProps[] =>
    users
      .map((user) => {
        const matchingUser = siteUsers.find(
          (siteUser) =>
            (siteUser.secondaryText || "").toLowerCase() ===
            (user.email || "").toLowerCase(),
        );

        return {
          key: String(user.id || user.email),
          text: user.name || matchingUser?.text || user.email || "",
          secondaryText: user.email || matchingUser?.secondaryText || "",
        };
      })
      .filter((persona) => !!persona.secondaryText);

  const mapPersonaToPeople = (
    items?: IPersonaProps[],
  ): IPeoplePickerDetails[] => {
    if (!Array.isArray(items)) return [];

    return items
      .map((item) => {
        const email = item?.secondaryText || "";
        const matchedSiteUser: any = siteUsers.find(
          (user) =>
            (user.secondaryText || "").toLowerCase() === email.toLowerCase(),
        );
        const parsedId = Number(matchedSiteUser?.key || item?.key);

        return {
          id: Number.isNaN(parsedId) ? 0 : parsedId,
          name: item?.text || "",
          email,
        };
      })
      .filter((user) => !!user.email);
  };

  return (
    <div>
      <ConfirmDialog />
      <div className={styles.header}>
        <div className={styles.title}>Target configurations</div>

        <Button
          label="Add target"
          icon="pi pi-plus"
          className={styles.addBtn}
          onClick={handleAdd}
        />
      </div>

      <DataTable
        value={targetConfigData}
        paginator={targetConfigData && targetConfigData?.length > 8}
        rows={8}
        emptyMessage="No data found"
      >
        <Column sortable field="Technology" header="Technology"></Column>
        <Column sortable field="Target" header="Target"></Column>
        <Column
          sortable
          field="ProjectManager"
          header="Project Manager"
          body={renderProjectManagerColumn}
        ></Column>
        <Column
          sortable
          field="DeliveryHead"
          header="Delivery Head"
          body={renderDeliveryHeadColumn}
        ></Column>
        <Column
          field="Action"
          header="Action"
          body={renderActionColumn}
        ></Column>
      </DataTable>

      {/* Dialog */}
      <Dialog
        header={isEdit ? "Edit Target" : "Add Target"}
        visible={showDialog}
        style={{ width: "52vw", maxWidth: "900px" }}
        onHide={() => setShowDialog(false)}
        className={styles.dialog}
      >
        <div className={styles.form}>
          {/* Technology */}
          <div className={styles.field}>
            <label>Technology</label>
            <Dropdown
              value={technologyOptions.find(
                (x) => x.name === formData.Technology,
              )}
              options={technologyOptions}
              optionLabel="name"
              placeholder="Select Technology"
              onChange={(e) => {
                const technology = e.value?.name || "";
                setFormData((prev) => ({
                  ...prev,
                  Technology: technology,
                }));
                if (technology.trim()) {
                  setFormErrors((prev) => ({ ...prev, Technology: false }));
                }
              }}
            />
            {formErrors.Technology && (
              <span className={styles.requiredText}>
                Technology is required
              </span>
            )}
          </div>

          {/* Target */}
          <div className={styles.field}>
            <label>Target</label>
            <InputText
              value={formData.Target}
              onChange={(e) => {
                const targetValue = e.target.value.replace(/[^0-9]/g, "");
                setFormData((prev) => ({
                  ...prev,
                  Target: targetValue,
                }));
                if (targetValue.trim()) {
                  setFormErrors((prev) => ({ ...prev, Target: false }));
                }
              }}
            />
            {formErrors.Target && (
              <span className={styles.requiredText}>Target is required</span>
            )}
          </div>

          {/* Project Manager */}
          <div className={styles.field}>
            <label>Project Manager</label>
            <NormalPeoplePicker
              selectedItems={mapToPersona(formData.ProjectManager)}
              onResolveSuggestions={filterSiteUsers}
              getTextFromItem={getTextFromItem}
              pickerSuggestionsProps={suggestionsProps}
              itemLimit={1}
              resolveDelay={1000}
              onChange={(items?: IPersonaProps[]) => {
                const selectedProjectManagers = mapPersonaToPeople(items);
                setFormData((prev) => ({
                  ...prev,
                  ProjectManager: selectedProjectManagers,
                }));
                if (selectedProjectManagers.length) {
                  setFormErrors((prev) => ({
                    ...prev,
                    ProjectManager: false,
                  }));
                }
              }}
            />
            {formErrors.ProjectManager && (
              <span className={styles.requiredText}>
                Project Manager is required
              </span>
            )}
          </div>

          {/* Delivery Head */}
          <div className={styles.field}>
            <label>Delivery Head</label>
            <NormalPeoplePicker
              selectedItems={mapToPersona(formData.DeliveryHead)}
              onResolveSuggestions={filterSiteUsers}
              getTextFromItem={getTextFromItem}
              pickerSuggestionsProps={suggestionsProps}
              itemLimit={1}
              resolveDelay={1000}
              onChange={(items?: IPersonaProps[]) => {
                const selectedDeliveryHeads = mapPersonaToPeople(items);
                setFormData((prev) => ({
                  ...prev,
                  DeliveryHead: selectedDeliveryHeads,
                }));
                if (selectedDeliveryHeads.length) {
                  setFormErrors((prev) => ({
                    ...prev,
                    DeliveryHead: false,
                  }));
                }
              }}
            />
            {formErrors.DeliveryHead && (
              <span className={styles.requiredText}>
                Delivery Head is required
              </span>
            )}
          </div>

          {/* Buttons */}
          <div className={styles.footer}>
            <Button
              label="Cancel"
              className={styles.cancelButton}
              onClick={() => setShowDialog(false)}
            />
            <Button
              label={isEdit ? "Update" : "Save"}
              onClick={handleSave}
              className={styles.saveButton}
            />
          </div>
        </div>
      </Dialog>
    </div>
  );
};

export default TargetConfig;
