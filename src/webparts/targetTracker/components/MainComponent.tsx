/* eslint-disable @typescript-eslint/no-var-requires */
/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable eqeqeq */
/* eslint-disable @typescript-eslint/no-use-before-define */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable prefer-const */
/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from "react";
import * as moment from "moment";
import styles from "./MainComponent.module.scss";
import "../../../External/Css/Styles.css";
import TargetConfig from "./TargetConfiguration/TargetConfig";
import TargetReports from "./TargetReports/TargetReports";

const MainComponent = (props: any) => {
  const [activeTab, setActiveTab] = React.useState<"config" | "reports">(
    "config",
  );

  // Financial Year Logic (Apr → Mar)
  const getFinancialYear = () => {
    const today = moment();
    const year = today.year();
    const month = today.month(); // 0 = Jan, 3 = Apr

    if (month >= 3) {
      // April or after
      return `FY ${year}-${year + 1}`;
    } else {
      return `FY ${year - 1}-${year}`;
    }
  };

  const getFinancialRange = () => {
    const today = moment();
    const year = today.year();
    const month = today.month();

    if (month >= 3) {
      return `Apr ${year} – Mar ${year + 1}`;
    } else {
      return `Apr ${year - 1} – Mar ${year}`;
    }
  };

  return (
    <div className={styles.pageShell}>
      <aside className={styles.sidePanel}>
        <div className={styles.brand}>
          <div className={styles.brandTitle}>Target Tracker</div>
          <div className={styles.brandSubtitle}>Performance workspace</div>
        </div>
        <div className={styles.tabs}>
          <div
            className={`${styles.tab} ${activeTab === "config" ? styles.active : ""}`}
            onClick={() => setActiveTab("config")}
          >
            Configurations
          </div>

          <div
            className={`${styles.tab} ${activeTab === "reports" ? styles.active : ""}`}
            onClick={() => setActiveTab("reports")}
          >
            Reports
          </div>
        </div>
      </aside>

      <main className={styles.mainArea}>
        {/* Financial Year Card */}
        <div className={styles.fyCard}>
          <div className={styles.fyLeft}>
            <span className={styles.icon}>🕒</span>
            <div>
              <div className={styles.fyLabel}>
                {activeTab === "config"
                  ? "Current financial year"
                  : "Reports Period"}
              </div>
              <div className={styles.fyTitle}>{getFinancialYear()}</div>
            </div>
          </div>
          <div className={styles.fyRange}>{getFinancialRange()}</div>
        </div>

        {/* Content */}
        <div className={styles.content}>
          {activeTab === "config" ? (
            <TargetConfig spfxContext={props.spfxContext} />
          ) : (
            <TargetReports spfxContext={props.spfxContext} />
          )}
        </div>
      </main>
    </div>
  );
};

export default MainComponent;
