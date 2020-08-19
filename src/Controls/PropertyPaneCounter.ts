import {
    PropertyPaneFieldType
} from '@microsoft/sp-property-pane';

export interface IPropertyPaneCounterProps
{
  label: string;
  initialValue: number;
  onPropertyChanged: (newValue: number) => void;
}

export function PropertyPaneCounter(targetProp: string, props: IPropertyPaneCounterProps) {
  return {
    type: PropertyPaneFieldType.Custom,
    targetProperty: targetProp,
    properties: {
      key: 'counter_' + new Date().getTime(),
      label: props.label,
      onRender: (domElement: HTMLElement, context? : any) : void => {
        if(domElement) {
          domElement.innerHTML = `<span>${ props.label } <br/>
            <input type="button" id="btnctr" value="${ props.initialValue }" />
          `;

          let btn = domElement.querySelector<HTMLInputElement>("#btnctr");

          btn.onclick = () => {
            let n = Math.random() *100;
            props.onPropertyChanged(n);
            btn.value = `${ n }`;
          }
        }
      }
    }
  }
}
