import * as React from 'react';
import styles from './DemoPivotControl.module.scss';
import { IDemoPivotControlProps } from './IDemoPivotControlProps';
import { IDemoPivotControlState} from './IDemoPivotControlState';
import { IconButton, IIconProps, Icon } from 'office-ui-fabric-react';
import { IStyleSet } from 'office-ui-fabric-react/lib/Styling';
import { Label, ILabelStyles } from 'office-ui-fabric-react/lib/Label';
import { ResizeGroup } from "office-ui-fabric-react/lib/ResizeGroup";
import {
  OverflowSet,
  IOverflowSetStyles,
} from "office-ui-fabric-react/lib/OverflowSet";
import { IContextualMenuItem } from "office-ui-fabric-react/lib/ContextualMenu";
import { mergeStyles, IButtonStyles,Link } from "office-ui-fabric-react";

export interface IOverflowData {
  primary: IContextualMenuItem[];
  overflow: IContextualMenuItem[];
  cacheKey?: string;
}

export default class DemoPivotControl extends React.Component<
  IDemoPivotControlProps,
  IDemoPivotControlState
> {
  constructor(props: IDemoPivotControlProps) {
    super(props);
    this.state = {
      selectedkey: "Item 0",
      overflowItems: [],
      width: 0,
      selectedvalue:"Item Pivot Value 0"
    };
    this.onRenderItem = this.onRenderItem.bind(this);
    this.generateData = this.generateData.bind(this);
   
  }
  public overflowSetStyles: Partial<IOverflowSetStyles> = { root: { height: 40 } };


public generateData = (
  count: number
 
): IOverflowData => {
  const icons = ["Add", "Share", "Upload"];
  const dataItems = [];
  let cacheKey = "";
  for (let index = 0; index < count; index++) {  
    const item = {
      key: `item${index}`,
      name: `Item ${index}`,
      value: `Item Pivot Value ${index}`,
      icon: icons[index % icons.length],
      onClick: (ev?: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>, item?: IContextualMenuItem) =>{
       this.setState({
         ...this.state,
         selectedkey: item.key,
         selectedvalue: item.value,
       });
    }
    };
    cacheKey = cacheKey + item.key;
    dataItems.push(item);
  }
  let result: IOverflowData = {
    primary: dataItems,
    overflow: [] as any[],
  };
  
  return result;
};



public onRenderItem = (item: any) => (                   
  <Link
    role="menuitem"
    styles={{ root: { marginRight: 10 } }}
     onClick={()=>{this.setState({selectedkey:item.name, selectedvalue:item.value})}}>   
     { item.name}    
     </Link>
);

  public render(): React.ReactElement<IDemoPivotControlProps> {   
     const numberOfItems = 20;     
     const dataToRender = this.generateData(
       numberOfItems
     );
      let itemclass = mergeStyles({
        color: "black",
      });

     const onReduceData = (currentData: any) => {
       if (currentData.primary.length === 0) {
         return undefined;
       }
       const overflow = [
         ...currentData.primary.slice(-1),
         ...currentData.overflow,
       ];
       const primary = currentData.primary.slice(0, -1);
       let cacheKey = undefined;
     
       return { primary, overflow, cacheKey };
     };

     const onGrowData = (currentData: any) => {
       if (currentData.overflow.length === 0) {
         return undefined;
       }
       const overflow = currentData.overflow.slice(1);
       const primary = [
         ...currentData.primary,
         ...currentData.overflow.slice(0, 1),
       ];
       let cacheKey = undefined;
      
       return { primary, overflow, cacheKey };
     };

    const onRenderOverflowButton = (
      overflowItems: any[] | undefined
    ): JSX.Element => {
      const buttonStyles: Partial<IButtonStyles> = {
        root: {
          minWidth: 0,
          padding: "0 4px",
          alignSelf: "stretch",
          height: "auto",
        },
      };
      return (
        <IconButton
          role="menuitem"
          title="More options"
          styles={buttonStyles}
          menuIconProps={{ iconName: "More" }}
          menuProps={{ items: overflowItems! }}
        />
      );
    };
     const onRenderData = (data: any) => {
       return (
       
           <OverflowSet
             role="menu"
             items={data.primary}
             overflowItems={data.overflow.length ? data.overflow : null}
             onRenderItem={this.onRenderItem}
             onRenderOverflowButton={onRenderOverflowButton}
             styles={this.overflowSetStyles}
           />
      
       );
     };

   
    return (
      <div className={styles.demoPivotControl}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <ResizeGroup
                role="tabpanel"
                aria-label="Resize Group with an Overflow Set"
                data={dataToRender}
                onReduceData={onReduceData}
                onGrowData={true ? onGrowData : undefined}
                onRenderData={onRenderData}
              />
              {this.state.selectedkey && (
                <div className={itemclass}>{this.state.selectedvalue}</div>
              )}
            </div>
          </div>
        </div>
      </div>
    );
  }
}
