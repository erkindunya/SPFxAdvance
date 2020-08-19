import { IPropertyPaneField, IPropertyPaneCustomFieldProps, 
    PropertyPaneFieldType } from "@microsoft/sp-property-pane";

interface IPropertyPaneTaxRateProps {
    label: string;
    selectedRate: number;
    onPropertyChanged: (propName:string, newValue: number) => void;
}

interface IPropertyPaneTaxRateInternalProps extends IPropertyPaneTaxRateProps,
    IPropertyPaneCustomFieldProps {

}

const taxRates : any[] = [
    {
        rate: 0,
        name: 'Zero Rated'
    },
    {
        rate: 5,
        name: '5% GST'
    },
    {
        rate: 8,
        name: '8% GST'
    },
    {
        rate: 10,
        name: '10% GST'
    },
    {
        rate: 12,
        name: '12% GST'
    },
    {
        rate: 18,
        name: '18% GST'
    }
];

export class PropertyPaneTaxRate implements IPropertyPaneField<IPropertyPaneTaxRateProps> {
    type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
    targetProperty: string;
    shouldFocus?: boolean;
    properties: IPropertyPaneTaxRateInternalProps;
    elem: HTMLElement;

    constructor(targetProp: string, props: IPropertyPaneTaxRateProps) {
        this.targetProperty = targetProp;
        this.shouldFocus = true;

        this.properties = {
            key: 'taxpicker_' + new Date().getTime(),
            label:props.label,
            selectedRate: props.selectedRate,
            onRender: this.onRender,
            onDispose:this.onDispose,
            onPropertyChanged: props.onPropertyChanged
        }
    }

    private getTaxRateDropDown() : string {
        let html = `${ this.properties.label } : <br/>
            <select id="${ this.properties.key }">
        `;

        for(let r of taxRates) {
            html += `<option ${ this.properties.selectedRate == r.rate? "selected" :""} value="${ r.rate }">${ r.name }</option>`;
        }

        return html + "</select>";
    }

    public onRender = (domElement:HTMLElement, context? : any) :void =>{
        this.elem = domElement;

        if(!this.elem) {
            console.log("PropertyPaneTaxRate.onRender()->Unable to find DOM Element");
            return;
        }

        this.elem.innerHTML = this.getTaxRateDropDown();

        let dd : HTMLSelectElement = this.elem.querySelector<HTMLSelectElement>(`#${ this.properties.key }`);

        dd.onchange = (event: Event) => {
            this.properties.selectedRate = parseFloat((<HTMLSelectElement>event.target).value);

            this.properties.onPropertyChanged(this.targetProperty,this.properties.selectedRate);
        }
    }

    public onDispose =  (domElement:HTMLElement, context? : any) :void =>{

    }
}