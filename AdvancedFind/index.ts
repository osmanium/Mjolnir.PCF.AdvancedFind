import { IInputs, IOutputs } from './generated/ManifestTypes';

import $ = require('jquery');

type DynamicIFrame = HTMLIFrameElement & {
  contentWindow: HTMLElement;
};
export class AdvancedFind implements ComponentFramework.StandardControl<IInputs, IOutputs> {
  private _advancedFindIFrame: HTMLElement;
  private _saveFetchXmlButton: HTMLElement;
  private _container: HTMLDivElement;
  private _controlViewRendered: boolean;
  private _fetchXml: string;
  private _entityLogicalName: string;
  private _objectTypeCode: number;
  private _setFetchXmlTimer: NodeJS.Timeout;
  private _notifyOutputChanged: () => void;
  private _context: ComponentFramework.Context<IInputs>;

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement,
  ): void {
    this._context = context;
    this._container = container;
    this._notifyOutputChanged = notifyOutputChanged;
    this._controlViewRendered = false;
    this._fetchXml = context.parameters.fetchXmlProperty.raw || '';
    this._entityLogicalName = context.parameters.entityLogicalNameProperty.raw || '';
    this._context.mode.trackContainerResize(false);
    console.log('init');
  }

  public async updateView(context: ComponentFramework.Context<IInputs>): Promise<void> {
    const newEntityLogicalName = context.parameters.entityLogicalNameProperty.raw || '';
    if (newEntityLogicalName !== this._entityLogicalName) {
      this._container.innerText = '';
      this._entityLogicalName = newEntityLogicalName;
      this._fetchXml = '';
      this._controlViewRendered = false;
    }

    if (!this._controlViewRendered) {
      console.log('updateView');

      this._controlViewRendered = true;
      await this.renderAdvancedFindIFrame();
    }
  }

  private async renderAdvancedFindIFrame(): Promise<void> {
    if (this._setFetchXmlTimer) clearInterval(this._setFetchXmlTimer);

    const metadata = await this._context.utils.getEntityMetadata(this._entityLogicalName);

    /* eslint-disable @typescript-eslint/restrict-template-expressions */
    /* eslint-disable @typescript-eslint/no-unsafe-member-access */
    /* eslint-disable @typescript-eslint/no-unsafe-call */
    /* eslint-disable @typescript-eslint/no-explicit-any */
    const frameSrc = `${(<any>(
      this._context
    )).page.getClientUrl()}/SFA/goal/ParticipatingQueryCondition.aspx?entitytypecode=${
      metadata.ObjectTypeCode
    }&readonlymode=false`;

    //Advanced Find IFrame
    this.AddAdvancedFindIFrame(frameSrc);
  }

  private AddAdvancedFindIFrame(frameSrc: string) {
    this._advancedFindIFrame = document.createElement('iframe');
    this._advancedFindIFrame.id = 'AdvancedFind_IFRAME';
    this._advancedFindIFrame.classList.add('AdvancedFind_IFrame');
    this._advancedFindIFrame.setAttribute('src', frameSrc);
    this._container.appendChild(this._advancedFindIFrame);
    this._advancedFindIFrame.onload = () => {
      //If FetchXml is already saved on the record, set in IFrame
      this._setFetchXmlTimer = global.setInterval(() => {
        if (
          <any>this._advancedFindIFrame !== null &&
          <any>this._advancedFindIFrame !== undefined &&
          (<any>this._advancedFindIFrame).contentWindow !== null &&
          (<any>this._advancedFindIFrame).contentWindow !== undefined &&
          (<any>this._advancedFindIFrame).contentWindow.advFind !== null &&
          (<any>this._advancedFindIFrame).contentWindow.advFind !== undefined
        ) {
          if (this._setFetchXmlTimer) {
            clearInterval(this._setFetchXmlTimer);
          }

          this.SetFetchXml();
          this.AddSaveButton();
        }
      }, 200);
    };
  }

  private AddSaveButton() {
    //Save Button
    this._saveFetchXmlButton = document.createElement('button');
    this._saveFetchXmlButton.id = 'getFetchXmlbtn';
    this._saveFetchXmlButton.innerText = 'Save';
    this._saveFetchXmlButton.classList.add('ms-crm-Button');
    this._saveFetchXmlButton.style.cssFloat = 'right';
    this._saveFetchXmlButton.setAttribute('onmouseover', 'Mscrm.ButtonUtils.hoverOn(this);');
    this._saveFetchXmlButton.setAttribute('onmouseout', 'Mscrm.ButtonUtils.hoverOff(this);');
    this._saveFetchXmlButton.onclick = () => {
      const newFetchXml: string = (<any>(
        this._advancedFindIFrame
      )).contentWindow.advFind.control.get_fetchXml() as string;

      if (newFetchXml !== this._fetchXml) {
        this._fetchXml = newFetchXml;
        this._notifyOutputChanged();
        (<any>window).Xrm.Page.data.save();
      }
    };

    $('#' + this._advancedFindIFrame.id)
      .contents()
      .find('#buttonPreview')
      .after(this._saveFetchXmlButton);
  }

  private SetFetchXml() {
    if (this._fetchXml !== '') {
      (<any>this._advancedFindIFrame).contentWindow.advFind.control.set_fetchXml(this._fetchXml);
    }
  }

  public getOutputs(): IOutputs {
    console.log('getOutputs');
    return {
      fetchXmlProperty: this._fetchXml,
    } as IOutputs;
  }

  // eslint-disable-next-line @typescript-eslint/no-empty-function
  public destroy(): void {}
}
