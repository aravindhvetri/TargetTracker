/* eslint-disable @typescript-eslint/no-var-requires */
/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable eqeqeq */
/* eslint-disable @typescript-eslint/no-use-before-define */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable prefer-const */
/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from "react";
import { useState, useEffect, useMemo } from "react";
import SPServices from "../../../../External/CommonServices/SPServices";
import { Config } from "../../../../External/CommonServices/Config";
import { DataTable } from "primereact/datatable";
import { Column } from "primereact/column";
import {
  filterSiteUsers,
  getCurrentFinancialYearOption,
  getFinancialYearFilterOptions,
  getTextFromItem,
  mapPersonaToPeople,
  mapToPersona,
  normalizeFinancialYear,
  peopleColumnTemplate,
  peoplePickerSuggestionsProps,
} from "../../../../External/CommonTemplates/CommonTemplates";
import styles from "./TargetConfiguration.module.scss";
import { Button } from "primereact/button";
import { Dialog } from "primereact/dialog";
import { ConfirmDialog, confirmDialog } from "primereact/confirmdialog";
import { InputText } from "primereact/inputtext";
import { Dropdown } from "primereact/dropdown";
import { IPersonaProps, NormalPeoplePicker } from "@fluentui/react";

type IFormErrors = {
  Technology: boolean;
  FinancialYear: boolean;
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
  const [fieldChoiceOptions, setFieldChoiceOptions] = useState<
    Record<string, { name: string }[]>
  >({
    Technology: [],
    FinancialYear: [],
  });
  const [siteUsers, setSiteUsers] = useState<IPersonaProps[]>([]);
  const [formErrors, setFormErrors] = useState<IFormErrors>({
    Technology: false,
    FinancialYear: false,
    Target: false,
    ProjectManager: false,
    DeliveryHead: false,
  });
  const [listSearch, setListSearch] = useState("");
  const [selectedFinancialYear, setSelectedFinancialYear] =
    useState<string>("");

  //Use Effect to get Financial Year Filter Options:
  useEffect(() => {
    const availableFinancialYears = getFinancialYearFilterOptions(
      fieldChoiceOptions.FinancialYear,
      targetConfigData,
    );
    if (!availableFinancialYears.length) {
      if (selectedFinancialYear) {
        setSelectedFinancialYear("");
      }
      return;
    }

    const selectedStillValid = availableFinancialYears.some(
      (option) => option.name === selectedFinancialYear,
    );
    if (selectedStillValid) return;

    const currentFinancialYear = getCurrentFinancialYearOption(
      availableFinancialYears,
    );
    setSelectedFinancialYear(
      currentFinancialYear || availableFinancialYears[0].name || "",
    );
  }, [fieldChoiceOptions.FinancialYear, targetConfigData]);

  //Filtered Target Configuration Data function:
  const filteredTargetConfigData = useMemo(() => {
    const q = listSearch.trim().toLowerCase();
    const normalizedSelectedFinancialYear = normalizeFinancialYear(
      selectedFinancialYear,
    );

    const peopleMatch = (people: IPeoplePickerDetails[]): boolean =>
      (people || []).some((p) => {
        const name = (p?.name || "").toLowerCase();
        const email = (p?.email || "").toLowerCase();
        return name.includes(q) || email.includes(q);
      });

    return targetConfigData.filter((row) => {
      const rowFinancialYear = normalizeFinancialYear(row.FinancialYear || "");
      if (
        normalizedSelectedFinancialYear &&
        rowFinancialYear !== normalizedSelectedFinancialYear
      ) {
        return false;
      }

      if (!q) return true;

      const technology = (row.Technology || "").toLowerCase();
      const financialYear = (row.FinancialYear || "").toLowerCase();
      const target = String(row.Target ?? "").toLowerCase();
      if (
        technology.includes(q) ||
        financialYear.includes(q) ||
        target.includes(q)
      )
        return true;
      if (peopleMatch(row.ProjectManager)) return true;
      if (peopleMatch(row.DeliveryHead)) return true;
      return false;
    });
  }, [targetConfigData, listSearch, selectedFinancialYear]);

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
            FinancialYear: items.FinancialYear || "",
            Target: items.Target || "",
            ProjectManager: _ProjectManager || [],
            DeliveryHead: _DeliveryHead || [],
            IsDelete: items.IsDelete || false,
          });
        });
        setTargetConfigData([...data]);
        getFieldChoices();
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

  //Get Choice column values function:
  const getFieldChoices = async () => {
    const fieldNames = ["Technology", "FinancialYear"];
    try {
      const response = await Promise.all(
        fieldNames.map((fieldName) =>
          SPServices.SPGetChoices({
            Listname: Config.ListNames.TargetConfiguration,
            FieldName: fieldName,
          }),
        ),
      );

      const nextChoiceOptions: Record<string, { name: string }[]> = {
        Technology: [],
        FinancialYear: [],
      };

      response.forEach((res: any, index: number) => {
        const fieldName = fieldNames[index];
        nextChoiceOptions[fieldName] = (res?.Choices || []).map((val: any) => ({
          name: val,
        }));
      });

      setFieldChoiceOptions(nextChoiceOptions);
    } catch (error) {
      console.log(error);
      setFieldChoiceOptions({
        Technology: [],
        FinancialYear: [],
      });
    }
  };

  //Render Manager Column function:
  const renderProjectManagerColumn = (rowData: ITargetConfigurationData) =>
    peopleColumnTemplate(rowData?.ProjectManager || []);

  //Render Delivery Head Column function:
  const renderDeliveryHeadColumn = (rowData: ITargetConfigurationData) =>
    peopleColumnTemplate(rowData?.DeliveryHead || []);

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
    const currentFinancialYear = getCurrentFinancialYearOption(
      fieldChoiceOptions.FinancialYear,
    );
    setFormData({
      ...Config.TargetConfigurationData,
      FinancialYear:
        currentFinancialYear || fieldChoiceOptions.FinancialYear[0]?.name || "",
    });
    setFormErrors({
      Technology: false,
      FinancialYear: false,
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
      FinancialYear: false,
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
      FinancialYear: !formData.FinancialYear?.trim(),
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
      FinancialYear: formData.FinancialYear || "",
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

  const financialYearFilterOptions = getFinancialYearFilterOptions(
    fieldChoiceOptions.FinancialYear,
    targetConfigData,
  );

  return (
    <div>
      <ConfirmDialog />
      <div className={styles.header}>
        <div className={styles.title}>Target configurations</div>

        <div className={styles.headerActions}>
          <div className={styles.searchField}>
            <i className={`pi pi-search ${styles.searchIcon}`} aria-hidden />
            <InputText
              value={listSearch}
              onChange={(e) => setListSearch(e.target.value)}
              placeholder="Search…"
              className={styles.listSearchInput}
              aria-label="Search target configurations"
            />
          </div>
          <div className={styles.financialYearFilter}>
            <Dropdown
              value={selectedFinancialYear}
              options={financialYearFilterOptions}
              optionLabel="name"
              optionValue="name"
              placeholder="Financial Year"
              className={styles.financialYearDropdown}
              onChange={(e) => setSelectedFinancialYear(e.value || "")}
              aria-label="Filter by financial year"
            />
          </div>
          <Button
            label="Add target"
            icon="pi pi-plus"
            className={styles.addBtn}
            onClick={handleAdd}
          />
        </div>
      </div>

      <DataTable
        value={filteredTargetConfigData}
        paginator={
          filteredTargetConfigData && filteredTargetConfigData?.length > 7
        }
        rows={7}
        emptyMessage="No data found"
      >
        <Column sortable field="FinancialYear" header="Financial Year"></Column>
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
              value={fieldChoiceOptions.Technology.find(
                (x) => x.name === formData.Technology,
              )}
              options={fieldChoiceOptions.Technology}
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

          {/* Financial Year */}
          <div className={styles.field}>
            <label>Financial Year</label>
            <Dropdown
              value={fieldChoiceOptions.FinancialYear.find(
                (x) => x.name === formData.FinancialYear,
              )}
              options={fieldChoiceOptions.FinancialYear}
              optionLabel="name"
              placeholder="Select Financial Year"
              onChange={(e) => {
                const financialYear = e.value?.name || "";
                setFormData((prev) => ({
                  ...prev,
                  FinancialYear: financialYear,
                }));
                if (financialYear.trim()) {
                  setFormErrors((prev) => ({
                    ...prev,
                    FinancialYear: false,
                  }));
                }
              }}
            />
            {formErrors.FinancialYear && (
              <span className={styles.requiredText}>
                Financial Year is required
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
              selectedItems={mapToPersona(formData.ProjectManager, siteUsers)}
              onResolveSuggestions={(filterText, currentPersonas) =>
                filterSiteUsers(filterText, currentPersonas, siteUsers)
              }
              getTextFromItem={getTextFromItem}
              pickerSuggestionsProps={peoplePickerSuggestionsProps}
              itemLimit={1}
              resolveDelay={1000}
              onChange={(items?: IPersonaProps[]) => {
                const selectedProjectManagers = mapPersonaToPeople(
                  items,
                  siteUsers,
                );
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
              selectedItems={mapToPersona(formData.DeliveryHead, siteUsers)}
              onResolveSuggestions={(filterText, currentPersonas) =>
                filterSiteUsers(filterText, currentPersonas, siteUsers)
              }
              getTextFromItem={getTextFromItem}
              pickerSuggestionsProps={peoplePickerSuggestionsProps}
              itemLimit={1}
              resolveDelay={1000}
              onChange={(items?: IPersonaProps[]) => {
                const selectedDeliveryHeads = mapPersonaToPeople(
                  items,
                  siteUsers,
                );
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
