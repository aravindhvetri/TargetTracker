import * as React from "react";
import {
  IBasePickerSuggestionsProps,
  IPersonaProps,
  PersonaSize,
} from "@fluentui/react";
import {
  DirectionalHint,
  Label,
  Persona,
  PersonaPresence,
  TooltipDelay,
  TooltipHost,
} from "office-ui-fabric-react";
import "../Css/Styles.css";

export const getCurrentFinancialYearLabel = (): string => {
  const today = new Date();
  const year = today.getFullYear();
  const month = today.getMonth() + 1;
  const fyStartYear = month >= 4 ? year : year - 1;
  const fyEndYear = fyStartYear + 1;
  return `${fyStartYear}-${fyEndYear}`;
};

export const normalizeFinancialYear = (value: string): string => {
  return (value || "").replace(/\s+/g, "").toLowerCase();
};

export const getFinancialYearFilterOptions = (
  financialYearChoices: { name: string }[],
  targetConfigData: ITargetConfigurationData[],
): { name: string }[] => {
  const combinedValues = [
    ...(financialYearChoices || []).map((option) => option.name),
    ...targetConfigData.map((row) => row.FinancialYear),
  ]
    .map((value) => (value || "").trim())
    .filter(Boolean);

  const uniqueValues = Array.from(new Set(combinedValues));
  return uniqueValues.map((name) => ({ name }));
};

export const getCurrentFinancialYearOption = (
  options: { name: string }[],
): string => {
  const currentFinancialYear = getCurrentFinancialYearLabel();
  const normalizedCurrentFinancialYear =
    normalizeFinancialYear(currentFinancialYear);

  const directMatch = options.find(
    (option) =>
      normalizeFinancialYear(option.name) === normalizedCurrentFinancialYear,
  );
  if (directMatch) return directMatch.name;

  const yearTokens = currentFinancialYear.split("-");
  const partialMatch = options.find((option) => {
    const normalizedOption = normalizeFinancialYear(option.name);
    return yearTokens.every((token) => normalizedOption.includes(token));
  });

  return partialMatch?.name || "";
};

export const getNumericValue = (value: any): number => {
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

export const getAmountFromBilling = (billing: any): number => {
  const amount = getNumericValue(billing?.Amount);
  if (amount > 0) return amount;

  const tmAmount = getNumericValue(billing?.TMAmount);
  if (tmAmount > 0) return tmAmount;

  return getNumericValue(billing?.MonthlyAmount);
};

export const getBillingProjectId = (billing: any): number => {
  return (
    billing?.CRMProjectId ||
    billing?.ProjectId ||
    billing?.CRMProject?.Id ||
    billing?.Project?.Id ||
    billing?.ProjectLookupId ||
    0
  );
};

export const formatCurrency = (value: number): string => {
  return `${Math.round(value).toLocaleString("en-IN")}`;
};

export const getFinancialYearDateRange = (
  financialYearLabel: string,
): { startFY: Date; endFY: Date } | null => {
  const fallbackYear = getCurrentFinancialYearLabel();
  const source = (financialYearLabel || fallbackYear).trim();
  const yearMatch = source.match(/\d{4}/g);

  if (!yearMatch?.length) return null;

  const startYear = Number(yearMatch[0]);
  if (Number.isNaN(startYear)) return null;

  return {
    startFY: new Date(startYear, 3, 1),
    endFY: new Date(startYear + 1, 2, 31),
  };
};

//MultiPeoplePicker Template:
export const multiPeoplePickerTemplate = (users: IPeoplePickerDetails[]) => {
  if (!users?.length) return null;

  const uniqueUsers = users.filter(
    (item, index, self) =>
      index === self.findIndex((t) => t?.email === item?.email),
  );

  return (
    <div
      className="user-selector-group"
      style={{
        display: "flex",
      }}
    >
      {uniqueUsers.map((value, index) => {
        if (index < 2) {
          return (
            <Persona
              key={index}
              styles={{
                root: {
                  cursor: "pointer",
                  margin: "0 !important",
                  ".ms-Persona-details": {
                    display: "none",
                  },
                },
              }}
              imageUrl={`/_layouts/15/userphoto.aspx?size=S&username=${value.email}`}
              title={value.name}
              size={PersonaSize.size24}
            />
          );
        }
        return null;
      })}

      {uniqueUsers.length > 2 && (
        <TooltipHost
          className="all-member-users"
          content={
            <ul style={{ margin: 10, padding: 0 }}>
              {uniqueUsers.map((DName: any, index) => (
                <li key={index} style={{ listStyleType: "none" }}>
                  <div style={{ display: "flex" }}>
                    <Persona
                      showOverflowTooltip
                      size={PersonaSize.size24}
                      presence={PersonaPresence.none}
                      showInitialsUntilImageLoads
                      imageUrl={`/_layouts/15/userphoto.aspx?size=S&username=${DName.email}`}
                    />
                    <Label style={{ marginLeft: 10, fontSize: 12 }}>
                      {DName.name}
                    </Label>
                  </div>
                </li>
              ))}
            </ul>
          }
          delay={TooltipDelay.zero}
          directionalHint={DirectionalHint.bottomCenter}
          styles={{ root: { display: "inline-block" } }}
        >
          <div className="persona">
            +{uniqueUsers.length - 2}
            <div className="allPersona"></div>
          </div>
        </TooltipHost>
      )}
    </div>
  );
};

//PeoplePicker Template:
export const peoplePickerTemplate = (user: IPeoplePickerDetails) => {
  return (
    <>
      {user && (
        <div
          style={{
            display: "flex",
            gap: "10px",
            alignItems: "center",
          }}
        >
          <Persona
            styles={{
              root: {
                margin: "0 !important;",
                ".ms-Persona-details": {
                  display: "none",
                },
              },
            }}
            imageUrl={
              "/_layouts/15/userphoto.aspx?size=S&username=" + user?.email
            }
            title={user?.name}
            size={PersonaSize.size24}
          />
          <p
            style={{
              overflow: "hidden",
              textOverflow: "ellipsis",
              margin: 0,
            }}
            className="displayText"
            title={user?.name}
          >
            {user?.name}
          </p>
        </div>
      )}
    </>
  );
};

export const peopleColumnTemplate = (users: IPeoplePickerDetails[]) => {
  return (
    <div>
      {users?.length > 1
        ? multiPeoplePickerTemplate(users)
        : peoplePickerTemplate(users?.[0])}
    </div>
  );
};

export const peoplePickerSuggestionsProps: IBasePickerSuggestionsProps = {
  suggestionsHeaderText: "Suggested users",
  noResultsFoundText: "No matching users",
};

export const filterSiteUsers = (
  filterText: string,
  currentPersonas: IPersonaProps[] | undefined,
  siteUsers: IPersonaProps[],
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

export const getTextFromItem = (persona: IPersonaProps): string =>
  persona.text || "";

export const mapToPersona = (
  users: IPeoplePickerDetails[],
  siteUsers: IPersonaProps[],
): IPersonaProps[] =>
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

export const mapPersonaToPeople = (
  items: IPersonaProps[] | undefined,
  siteUsers: IPersonaProps[],
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
