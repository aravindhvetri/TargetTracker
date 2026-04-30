import * as React from "react";
import { PersonaSize } from "@fluentui/react";
import {
  DirectionalHint,
  Label,
  Persona,
  PersonaPresence,
  TooltipDelay,
  TooltipHost,
} from "office-ui-fabric-react";
import "../Css/Styles.css";

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
