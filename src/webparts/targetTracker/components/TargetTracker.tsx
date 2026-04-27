import * as React from "react";
import type { ITargetTrackerProps } from "./ITargetTrackerProps";
import { sp } from "@pnp/sp";
import { graph } from "@pnp/pnpjs";
import MainComponent from "./MainComponent";

export default class TargetTracker extends React.Component<
  ITargetTrackerProps,
  {}
> {
  constructor(prop: ITargetTrackerProps, state: {}) {
    super(prop);
    sp.setup({
      spfxContext: this.props.context as unknown as undefined,
    });

    graph.setup({
      spfxContext: this.props.context as unknown as undefined,
    });
  }

  public render(): React.ReactElement<ITargetTrackerProps> {
    return <MainComponent spfxContext={this.props.context} spContext={sp} />;
  }
}
