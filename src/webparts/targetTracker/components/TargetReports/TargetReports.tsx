/* eslint-disable @typescript-eslint/no-var-requires */
/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable eqeqeq */
/* eslint-disable @typescript-eslint/no-use-before-define */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable prefer-const */
/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from "react";
import styles from "./TargetReports.module.scss";
import { useState, useEffect, useMemo } from "react";
import SPServices from "../../../../External/CommonServices/SPServices";
import { Config } from "../../../../External/CommonServices/Config";
import { Dropdown } from "primereact/dropdown";
import { DataTable } from "primereact/datatable";
import { Column } from "primereact/column";
import { Dialog } from "primereact/dialog";
import { Button } from "primereact/button";
import {
  formatCurrency,
  getAmountFromBilling,
  getBillingProjectId,
  getCurrentFinancialYearOption,
  getFinancialYearDateRange,
  getFinancialYearFilterOptions,
  getNumericValue,
  normalizeFinancialYear,
} from "../../../../External/CommonTemplates/CommonTemplates";

interface IDeliveryHeadOption {
  label: string;
  value: IPeoplePickerDetails;
}

interface IProjectManagerReport {
  email: string;
  name: string;
  target: number;
  achieved: number;
  invoiceRaised: number;
  invoiceNotGenerated: number;
  variance: number;
  attainment: number;
}

interface IInvoiceDetailRow {
  id: number;
  projectName: string;
  clientName: string;
  status: number;
  amount: number;
  dueDate?: string;
  invoiceID?: string;
}

const TargetReports = (props: any) => {
  //States:
  const [targetConfigData, setTargetConfigData] = useState<
    ITargetConfigurationData[]
  >([]);
  const [deliveryHeadOptions, setDeliveryHeadOptions] = useState<
    IDeliveryHeadOption[]
  >([]);
  const [selectedDeliveryHead, setSelectedDeliveryHead] = useState<any>(null);
  const [fieldChoiceOptions, setFieldChoiceOptions] = useState<
    Record<string, { name: string }[]>
  >({
    FinancialYear: [],
  });
  const [selectedFinancialYear, setSelectedFinancialYear] =
    useState<string>("");
  const [projectManagerReports, setProjectManagerReports] = useState<
    IProjectManagerReport[]
  >([]);
  const [loading, setLoading] = useState<boolean>(false);
  const [invoiceDialogVisible, setInvoiceDialogVisible] =
    useState<boolean>(false);
  const [selectedInvoicePm, setSelectedInvoicePm] =
    useState<IProjectManagerReport | null>(null);
  const [invoiceDetailsByPM, setInvoiceDetailsByPM] = useState<
    Record<string, IInvoiceDetailRow[]>
  >({});

  //Use Effect to get Target Configuration Data:
  useEffect(() => {
    getTargetConfigurationData();
  }, []);

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

  const financialYearFilterOptions = getFinancialYearFilterOptions(
    fieldChoiceOptions.FinancialYear,
    targetConfigData,
  );

  const filteredDeliveryHeadOptions = useMemo(() => {
    const normalizedSelectedFinancialYear = normalizeFinancialYear(
      selectedFinancialYear,
    );
    const allDeliveryHeads: IPeoplePickerDetails[] = [];

    targetConfigData.forEach((item) => {
      const rowFinancialYear = normalizeFinancialYear(item.FinancialYear || "");
      if (
        normalizedSelectedFinancialYear &&
        rowFinancialYear !== normalizedSelectedFinancialYear
      ) {
        return;
      }

      if (item.DeliveryHead?.length) {
        item.DeliveryHead.forEach((dh) => {
          allDeliveryHeads.push(dh);
        });
      }
    });

    const uniqueDeliveryHeads = Array.from(
      new Map(allDeliveryHeads.map((item) => [item.email, item])).values(),
    );

    return uniqueDeliveryHeads.map((dh) => ({
      label: dh.name,
      value: dh,
    }));
  }, [targetConfigData, selectedFinancialYear]);

  useEffect(() => {
    setDeliveryHeadOptions(filteredDeliveryHeadOptions);

    if (!selectedDeliveryHead) return;

    const selectedStillValid = filteredDeliveryHeadOptions.some(
      (option) => option.value?.email === selectedDeliveryHead?.email,
    );

    if (!selectedStillValid) {
      setSelectedDeliveryHead(null);
      setProjectManagerReports([]);
    }
  }, [filteredDeliveryHeadOptions, selectedDeliveryHead]);

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

  const getFieldChoices = async () => {
    try {
      const response: any = await SPServices.SPGetChoices({
        Listname: Config.ListNames.TargetConfiguration,
        FieldName: "FinancialYear",
      });

      setFieldChoiceOptions({
        FinancialYear: (response?.Choices || []).map((val: any) => ({
          name: val,
        })),
      });
    } catch (error) {
      console.log(error);
      setFieldChoiceOptions({ FinancialYear: [] });
    }
  };

  const getBadgeClass = (value: number): string => {
    return value >= 100 ? styles.badgeSuccess : styles.badgeWarning;
  };

  const getProgressClass = (row: IProjectManagerReport): string => {
    if (row.attainment >= 100) return styles.progressGreen;
    if (row.attainment >= 75) return styles.progressAmber;
    return styles.progressRed;
  };

  //Handle Delivery Head Change function:
  const handleDeliveryHeadChange = async (selectedDH: any) => {
    setSelectedDeliveryHead(selectedDH);
    setProjectManagerReports([]);
    setLoading(true);

    try {
      const normalizedSelectedFinancialYear = normalizeFinancialYear(
        selectedFinancialYear,
      );

      // 1) Collect PMs and targets from target configuration under selected DH.
      const projectManagersMap = new Map<
        string,
        { email: string; name: string; target: number }
      >();

      targetConfigData.forEach((item) => {
        const rowFinancialYear = normalizeFinancialYear(
          item.FinancialYear || "",
        );
        if (
          normalizedSelectedFinancialYear &&
          rowFinancialYear !== normalizedSelectedFinancialYear
        ) {
          return;
        }

        const isMappedToDH = item.DeliveryHead?.some(
          (dh) => dh.email === selectedDH.email,
        );

        if (!isMappedToDH) return;

        const targetValue = getNumericValue(item.Target);

        item.ProjectManager.forEach((pm) => {
          if (!pm?.email) return;

          const existing = projectManagersMap.get(pm.email);
          if (existing) {
            existing.target += targetValue;
            if (!existing.name && pm.name) {
              existing.name = pm.name;
            }
          } else {
            projectManagersMap.set(pm.email, {
              email: pm.email,
              name: pm.name || pm.email,
              target: targetValue,
            });
          }
        });
      });

      const projectManagerEmails = Array.from(projectManagersMap.keys());

      if (!projectManagerEmails.length) {
        setLoading(false);
        return;
      }

      // 2) Read CRM Projects and map project ID -> project manager email.
      const projectsResponse: any[] = await SPServices.SPReadItems({
        Listname: Config.ListNames.CRMProjects,
        Select:
          "*,ProjectManager/EMail,ProjectManager/Title,DeliveryHead/EMail",
        Expand: "ProjectManager,DeliveryHead",
      });

      const financialYearRange = getFinancialYearDateRange(
        selectedFinancialYear,
      );
      if (!financialYearRange) {
        setProjectManagerReports([]);
        setLoading(false);
        return;
      }
      const { startFY, endFY } = financialYearRange;

      const projectToProjectManagers = new Map<number, string[]>();
      const projectLookup = new Map<
        number,
        { projectName: string; clientName: string }
      >();

      projectsResponse.forEach((project) => {
        const dhMatch = project?.DeliveryHead?.some(
          (dh: any) => dh?.EMail === selectedDH.email,
        );
        if (!dhMatch) return;

        const projectPMs = (project?.ProjectManager || [])
          .map((pm: any) => pm?.EMail)
          .filter((email: string) => projectManagerEmails.includes(email));
        if (!projectPMs.length) return;

        const startDate = project?.StartDate
          ? new Date(project.StartDate)
          : null;
        const endDate = project?.PlannedEndDate
          ? new Date(project.PlannedEndDate)
          : null;
        if (!startDate || !endDate) return;
        if (startDate < startFY || endDate > endFY) return;

        projectToProjectManagers.set(project.ID, projectPMs);

        const projectName =
          project?.ProjectName || project?.Title || project?.Project || "";
        const clientName =
          project?.ClientName ||
          project?.Client?.Title ||
          project?.Client ||
          "";

        projectLookup.set(project.ID, {
          projectName,
          clientName,
        });
      });

      if (!projectToProjectManagers.size) {
        const noProjectReports: IProjectManagerReport[] =
          projectManagerEmails.map((pmEmail) => {
            const pm = projectManagersMap.get(pmEmail)!;
            return {
              email: pm.email,
              name: pm.name,
              target: pm.target,
              achieved: 0,
              invoiceRaised: 0,
              invoiceNotGenerated: 0,
              variance: -pm.target,
              attainment: 0,
            };
          });
        setProjectManagerReports(noProjectReports);
        setLoading(false);
        return;
      }

      // 3) Read billings and aggregate per PM based on project mapping.
      const billingsResponse: any[] = await SPServices.SPReadItems({
        Listname: Config.ListNames.CRMBillings,
        Select: "*",
      });

      const summaryByPM = new Map<
        string,
        { achieved: number; invoiceRaised: number; invoiceNotGenerated: number }
      >();
      const invoiceDetails: Record<string, IInvoiceDetailRow[]> = {};

      billingsResponse.forEach((billing) => {
        const projectId = getBillingProjectId(billing);
        if (!projectId) return;

        const pmEmails = projectToProjectManagers.get(projectId);
        if (!pmEmails?.length) return;

        const amount = getAmountFromBilling(billing);
        const status = Number(billing?.Status);

        pmEmails.forEach((pmEmail) => {
          if (!summaryByPM.has(pmEmail)) {
            summaryByPM.set(pmEmail, {
              achieved: 0,
              invoiceRaised: 0,
              invoiceNotGenerated: 0,
            });
          }

          const current = summaryByPM.get(pmEmail)!;

          if (status === 3) {
            current.achieved += amount;
          } else if (status === 1) {
            current.invoiceRaised += amount;
          } else if (status === 0) {
            current.invoiceNotGenerated += amount;
          }

          if (!invoiceDetails[pmEmail]) {
            invoiceDetails[pmEmail] = [];
          }

          const projectMeta = projectLookup.get(projectId) || {
            projectName: "",
            clientName: "",
          };

          invoiceDetails[pmEmail].push({
            id: billing.ID || 0,
            projectName: projectMeta.projectName,
            clientName: projectMeta.clientName,
            status,
            amount,
            dueDate: billing?.DueDate || "",
            invoiceID: billing?.InvoiceID || "",
          });
        });
      });

      const pmReports: IProjectManagerReport[] = projectManagerEmails.map(
        (pmEmail) => {
          const pmTarget = projectManagersMap.get(pmEmail)!;
          const pmSummary = summaryByPM.get(pmEmail) || {
            achieved: 0,
            invoiceRaised: 0,
            invoiceNotGenerated: 0,
          };
          const target = pmTarget.target;
          const achieved = pmSummary.achieved;
          const variance = achieved - target;
          const attainment = target > 0 ? (achieved / target) * 100 : 0;

          return {
            email: pmTarget.email,
            name: pmTarget.name,
            target,
            achieved,
            invoiceRaised: pmSummary.invoiceRaised,
            invoiceNotGenerated: pmSummary.invoiceNotGenerated,
            variance,
            attainment,
          };
        },
      );

      setProjectManagerReports(pmReports);
      setInvoiceDetailsByPM(invoiceDetails);
    } catch (error) {
      console.log(error);
    } finally {
      setLoading(false);
    }
  };

  const handleViewInvoiceDetails = (pmReport: IProjectManagerReport) => {
    setSelectedInvoicePm(pmReport);
    setInvoiceDialogVisible(true);
  };

  const getStatusLabel = (status: number): string => {
    if (status === 3) return "Paid";
    if (status === 1) return "Invoice Raised";
    if (status === 0) return "Not Generated";
    return String(status ?? "");
  };

  const statusBodyTemplate = (rowData: IInvoiceDetailRow) => {
    const label = getStatusLabel(rowData.status);
    if (rowData.status === 3) {
      return <span className={styles.statusPaid}>{label}</span>;
    }
    if (rowData.status === 1) {
      return <span className={styles.statusRaised}>{label}</span>;
    }
    if (rowData.status === 0) {
      return <span className={styles.statusNotGenerated}>{label}</span>;
    }
    return <span>{label}</span>;
  };

  const amountBodyTemplate = (rowData: IInvoiceDetailRow) =>
    formatCurrency(rowData.amount || 0);

  const dateBodyTemplate = (rowData: IInvoiceDetailRow) => {
    if (!rowData.dueDate) return "-";
    try {
      const parsed = new Date(rowData.dueDate);
      if (Number.isNaN(parsed.getTime())) return rowData.dueDate;
      return parsed.toLocaleDateString("en-IN");
    } catch {
      return rowData.dueDate;
    }
  };

  const totalTarget = projectManagerReports.reduce(
    (sum, item) => sum + item.target,
    0,
  );
  const totalAchieved = projectManagerReports.reduce(
    (sum, item) => sum + item.achieved,
    0,
  );
  const totalVariance = totalAchieved - totalTarget;
  const overallAttainment =
    totalTarget > 0 ? (totalAchieved / totalTarget) * 100 : 0;

  const selectedDeliveryHeadLabel =
    selectedDeliveryHead?.name || "Delivery head";

  const selectedPmInvoiceRows: IInvoiceDetailRow[] = selectedInvoicePm
    ? invoiceDetailsByPM[selectedInvoicePm.email] || []
    : [];

  const totalReceivedForSelectedPm = selectedPmInvoiceRows
    .filter((row) => row.status === 3)
    .reduce((sum, row) => sum + (row.amount || 0), 0);

  return (
    <div className={styles.container}>
      <div className={styles.filterRow}>
        <div style={{ minWidth: "220px", flex: "1 1 240px" }}>
          <div className={styles.filterLabel}>Financial year</div>
          <Dropdown
            value={selectedFinancialYear}
            options={financialYearFilterOptions}
            optionLabel="name"
            optionValue="name"
            onChange={(e) => {
              setSelectedFinancialYear(e.value || "");
              setSelectedDeliveryHead(null);
              setProjectManagerReports([]);
            }}
            placeholder="Select Financial Year"
            className={styles.dropdown}
          />
        </div>
        <div style={{ minWidth: "220px", flex: "1 1 240px" }}>
          <div className={styles.filterLabel}>Filter by delivery head</div>
          <Dropdown
            value={selectedDeliveryHead}
            options={deliveryHeadOptions}
            onChange={(e) => handleDeliveryHeadChange(e.value)}
            placeholder="Select Delivery Head"
            className={styles.dropdown}
          />
        </div>
      </div>

      {selectedDeliveryHead && (
        <>
          <div className={styles.summaryCards}>
            <div className={styles.summaryCard}>
              <div className={styles.cardTitle}>Total target</div>
              <div className={styles.cardValue}>
                {formatCurrency(totalTarget)}
              </div>
              <div className={styles.cardSubtext}>
                {projectManagerReports.length} PMs configured
              </div>
            </div>

            <div className={styles.summaryCard}>
              <div className={styles.cardTitle}>Total achieved</div>
              <div className={`${styles.cardValue} ${styles.positiveText}`}>
                {formatCurrency(totalAchieved)}
              </div>
              <div className={styles.cardSubtext}>
                {Math.round(overallAttainment)}% attainment
              </div>
            </div>

            <div className={styles.summaryCard}>
              <div className={styles.cardTitle}>Variance</div>
              <div
                className={`${styles.cardValue} ${
                  totalVariance >= 0 ? styles.positiveText : styles.negativeText
                }`}
              >
                {totalVariance >= 0
                  ? `+${formatCurrency(totalVariance)}`
                  : `-${formatCurrency(Math.abs(totalVariance))}`}
              </div>
              <div className={styles.cardSubtext}>
                {totalVariance >= 0 ? "Above target" : "Below target"}
              </div>
            </div>
          </div>

          <div className={styles.sectionTitle}>Project manager performance</div>
          {!loading && !projectManagerReports.length && (
            <div className={styles.emptyState}>
              No report data found for this delivery head.
            </div>
          )}
          {loading && (
            <div className={styles.emptyState}>Loading report data...</div>
          )}

          <div className={styles.pmCardsContainer}>
            {!loading &&
              projectManagerReports.map((pmReport) => (
                <div className={styles.pmCard} key={pmReport.email}>
                  <div className={styles.pmHeader}>
                    <div className={styles.pmIdentity}>
                      <div className={styles.pmAvatar}>
                        {(pmReport.name || "NA")
                          .split(" ")
                          .slice(0, 2)
                          .map((token) => token[0])
                          .join("")
                          .toUpperCase()}
                      </div>
                      <div>
                        <div className={styles.pmName}>{pmReport.name}</div>
                        <div className={styles.pmMeta}>
                          Under {selectedDeliveryHeadLabel} - Revenue
                        </div>
                      </div>
                    </div>
                    <div className={styles.pmTargetInfo}>
                      <div className={styles.pmAchieved}>
                        {formatCurrency(pmReport.achieved)}
                      </div>
                      <div className={styles.pmOfTarget}>
                        of {formatCurrency(pmReport.target)}
                      </div>
                    </div>
                  </div>

                  <div className={styles.progressTrack}>
                    <div
                      className={`${styles.progressFill} ${getProgressClass(pmReport)}`}
                      style={{
                        width: `${Math.min(pmReport.attainment, 100)}%`,
                      }}
                    />
                  </div>

                  <div className={styles.pmFooter}>
                    <div
                      className={`${styles.attainmentBadge} ${getBadgeClass(pmReport.attainment)}`}
                    >
                      {Math.round(pmReport.attainment)}%
                    </div>
                    <div className={styles.pmMetric}>
                      Target <span>{formatCurrency(pmReport.target)}</span>
                    </div>
                    <div className={styles.pmMetric}>
                      Achieved <span>{formatCurrency(pmReport.achieved)}</span>
                    </div>
                    <div className={styles.pmMetric}>
                      Invoice Raised{" "}
                      <span>{formatCurrency(pmReport.invoiceRaised)}</span>
                    </div>
                    <div className={styles.pmMetric}>
                      Not Generated{" "}
                      <span>
                        {formatCurrency(pmReport.invoiceNotGenerated)}
                      </span>
                    </div>
                    <div className={styles.pmMetric}>
                      Variance{" "}
                      <span
                        className={
                          pmReport.variance >= 0
                            ? styles.positiveText
                            : styles.negativeText
                        }
                      >
                        {pmReport.variance >= 0
                          ? `+${formatCurrency(pmReport.variance)}`
                          : `-${formatCurrency(Math.abs(pmReport.variance))}`}
                      </span>
                    </div>
                    <div className={styles.pmMetric}>
                      <Button
                        type="button"
                        label="View Invoice Details"
                        icon="pi pi-eye"
                        className={styles.viewInvoiceButton}
                        onClick={() => handleViewInvoiceDetails(pmReport)}
                      />
                    </div>
                  </div>
                </div>
              ))}
          </div>

          {!!projectManagerReports.length && (
            <>
              <div className={styles.sectionTitle}>Delivery head summary</div>
              <div className={styles.summaryTable}>
                <div className={styles.summaryHeader}>Delivery Head</div>
                <div className={styles.summaryHeader}>Total Target</div>
                <div className={styles.summaryHeader}>Achieved</div>
                <div className={styles.summaryHeader}>Variance</div>
                <div className={styles.summaryHeader}>Attainment</div>

                <div className={styles.summaryCell}>
                  {selectedDeliveryHeadLabel}
                </div>
                <div className={styles.summaryCell}>
                  {formatCurrency(totalTarget)}
                </div>
                <div className={styles.summaryCell}>
                  {formatCurrency(totalAchieved)}
                </div>
                <div
                  className={`${styles.summaryCell} ${
                    totalVariance >= 0
                      ? styles.positiveText
                      : styles.negativeText
                  }`}
                >
                  {totalVariance >= 0
                    ? `+${formatCurrency(totalVariance)}`
                    : `-${formatCurrency(Math.abs(totalVariance))}`}
                </div>
                <div className={styles.summaryCell}>
                  <span
                    className={`${styles.attainmentBadge} ${getBadgeClass(overallAttainment)}`}
                  >
                    {Math.round(overallAttainment)}%
                  </span>
                </div>
              </div>
            </>
          )}
        </>
      )}
      <Dialog
        header={
          selectedInvoicePm
            ? `Invoice details - ${selectedInvoicePm.name}`
            : "Invoice details"
        }
        visible={invoiceDialogVisible}
        style={{ width: "60vw" }}
        modal
        onHide={() => setInvoiceDialogVisible(false)}
        className={styles.invoiceDialog}
      >
        {selectedPmInvoiceRows.length ? (
          <>
            <div className={styles.invoiceHeaderMeta}>
              <div>
                <div className={styles.invoiceMetaLabel}>Project manager</div>
                <div className={styles.invoiceMetaValue}>
                  {selectedInvoicePm?.name}
                </div>
              </div>
            </div>

            <DataTable
              value={selectedPmInvoiceRows}
              paginator={selectedPmInvoiceRows.length > 8}
              rows={8}
              emptyMessage="No billing data available"
              className={styles.invoiceTable}
            >
              <Column field="projectName" header="Project Name" sortable />
              <Column field="clientName" header="Client Name" sortable />
              <Column
                field="amount"
                header="Amount"
                body={amountBodyTemplate}
                sortable
              />
              <Column
                field="status"
                header="Status"
                body={statusBodyTemplate}
                sortable
              />
              <Column field="invoiceID" header="Invoice ID" sortable />
              <Column
                field="dueDate"
                header="Invoice Raised Date"
                body={dateBodyTemplate}
                sortable
              />
            </DataTable>

            <div className={styles.invoiceFooter}>
              <div className={styles.invoiceFooterLabel}>
                Total received (Paid)
              </div>
              <div className={styles.invoiceFooterValue}>
                {formatCurrency(totalReceivedForSelectedPm)}
              </div>
            </div>
          </>
        ) : (
          <div className={styles.emptyState}>
            No billing data available for this project manager.
          </div>
        )}
      </Dialog>
    </div>
  );
};

export default TargetReports;
