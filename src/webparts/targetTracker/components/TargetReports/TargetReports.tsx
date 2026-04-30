/* eslint-disable @typescript-eslint/no-var-requires */
/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable eqeqeq */
/* eslint-disable @typescript-eslint/no-use-before-define */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable prefer-const */
/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from "react";
import styles from "./TargetReports.module.scss";
import { useState, useEffect } from "react";
import SPServices from "../../../../External/CommonServices/SPServices";
import { Config } from "../../../../External/CommonServices/Config";
import { Dropdown } from "primereact/dropdown";

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

const TargetReports = (props: any) => {
  //States:
  const [targetConfigData, setTargetConfigData] = useState<
    ITargetConfigurationData[]
  >([]);
  const [deliveryHeadOptions, setDeliveryHeadOptions] = useState<
    IDeliveryHeadOption[]
  >([]);
  const [selectedDeliveryHead, setSelectedDeliveryHead] = useState<any>(null);
  const [projectManagerReports, setProjectManagerReports] = useState<
    IProjectManagerReport[]
  >([]);
  const [loading, setLoading] = useState<boolean>(false);

  //Use Effect to get Target Configuration Data:
  useEffect(() => {
    getTargetConfigurationData();
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
        let allDeliveryHeads: any[] = [];

        // Get all delivery heads
        data.forEach((item) => {
          if (item.DeliveryHead?.length) {
            item.DeliveryHead.forEach((dh) => {
              allDeliveryHeads.push(dh);
            });
          }
        });

        // Remove duplicates using Map (based on email)
        const uniqueDeliveryHeads = Array.from(
          new Map(allDeliveryHeads.map((item) => [item.email, item])).values(),
        );

        // Convert to dropdown format
        const dropdownOptions = uniqueDeliveryHeads.map((dh) => ({
          label: dh.name,
          value: dh,
        }));

        setDeliveryHeadOptions(dropdownOptions);
      })
      .catch((error: any) => {
        console.log(error);
      });
  };

  //Handle Delivery Head Change function:
  const getNumericValue = (value: any): number => {
    if (value === null || value === undefined || value === "") {
      return 0;
    }
    if (typeof value === "number") {
      return Number.isNaN(value) ? 0 : value;
    }
    const cleanedValue = value.toString().replace(/,/g, "").trim();
    const parsed = parseFloat(cleanedValue);
    return Number.isNaN(parsed) ? 0 : parsed;
  };

  const getAmountFromBilling = (billing: any): number => {
    const amount = getNumericValue(billing?.Amount);
    if (amount > 0) return amount;

    const tmAmount = getNumericValue(billing?.TMAmount);
    if (tmAmount > 0) return tmAmount;

    return getNumericValue(billing?.MonthlyAmount);
  };

  const getBillingProjectId = (billing: any): number => {
    return (
      billing?.CRMProjectId ||
      billing?.ProjectId ||
      billing?.CRMProject?.Id ||
      billing?.Project?.Id ||
      billing?.ProjectLookupId ||
      0
    );
  };

  const formatCurrency = (value: number): string => {
    return `₹${Math.round(value).toLocaleString("en-IN")}`;
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
      // 1) Collect PMs and targets from target configuration under selected DH.
      const projectManagersMap = new Map<
        string,
        { email: string; name: string; target: number }
      >();

      targetConfigData.forEach((item) => {
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

      const today = new Date();
      const currentYear = today.getFullYear();
      const currentMonth = today.getMonth() + 1;

      const startFY =
        currentMonth >= 4
          ? new Date(currentYear, 3, 1)
          : new Date(currentYear - 1, 3, 1);
      const endFY =
        currentMonth >= 4
          ? new Date(currentYear + 1, 2, 31)
          : new Date(currentYear, 2, 31);

      const projectToProjectManagers = new Map<number, string[]>();

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
    } catch (error) {
      console.log(error);
    } finally {
      setLoading(false);
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

  return (
    <div className={styles.container}>
      <div className={styles.filterRow}>
        <div className={styles.filterLabel}>Filter by delivery head</div>
        <Dropdown
          value={selectedDeliveryHead}
          options={deliveryHeadOptions}
          onChange={(e) => handleDeliveryHeadChange(e.value)}
          placeholder="Select Delivery Head"
          className={styles.dropdown}
        />
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
    </div>
  );
};

export default TargetReports;
