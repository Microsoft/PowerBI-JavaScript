import { IHttpPostMessageResponse } from 'http-post-message';
import { IFilterable } from './ifilterable';
import { IReportNode } from './report';
import { VisualDescriptor } from './visualDescriptor';
import * as models from 'powerbi-models';
import * as utils from './util';
import * as errors from './errors';

/**
 * A Page node within a report hierarchy
 *
 * @export
 * @interface IPageNode
 */
export interface IPageNode {
  report: IReportNode;
  name: string;
}

/**
 * A Power BI report page
 *
 * @export
 * @class Page
 * @implements {IPageNode}
 * @implements {IFilterable}
 */
export class Page implements IPageNode, IFilterable {
  /**
   * The parent Power BI report that this page is a member of
   *
   * @type {IReportNode}
   */
  report: IReportNode;
  /**
   * The report page name
   *
   * @type {string}
   */
  name: string;

  /**
   * The user defined display name of the report page, which is undefined if the page is created manually
   *
   * @type {string}
   */
  displayName: string;

  /**
   * Is this page is the active page
   *
   * @type {boolean}
   */
  isActive: boolean;

  /**
   * The visibility of the page.
   * 0 - Always Visible
   * 1 - Hidden in View Mode
   *
   * @type {models.SectionVisibility}
   */
   visibility: models.SectionVisibility;

  /**
   * Page size as saved in the report.
   * @type {models.ICustomPageSize}
   */
  defaultSize: models.ICustomPageSize;

  /**
   * Page display options as saved in the report.
   * @type {models.ICustomPageSize}
   */
  defaultDisplayOption: models.DisplayOption;

  /**
   * Creates an instance of a Power BI report page.
   *
   * @param {IReportNode} report
   * @param {string} name
   * @param {string} [displayName]
   * @param {boolean} [isActivePage]
   * @param {models.SectionVisibility} [visibility]
   * @hidden
   */
  constructor(report: IReportNode, name: string, displayName?: string, isActivePage?: boolean, visibility?: models.SectionVisibility, defaultSize?: models.ICustomPageSize, defaultDisplayOption?: models.DisplayOption) {
    this.report = report;
    this.name = name;
    this.displayName = displayName;
    this.isActive = isActivePage;
    this.visibility = visibility;
    this.defaultSize = defaultSize;
    this.defaultDisplayOption = defaultDisplayOption;
  }

  /**
   * Gets all page level filters within the report.
   *
   * ```javascript
   * page.getFilters()
   *  .then(filters => { ... });
   * ```
   *
   * @returns {(Promise<models.IFilter[]>)}
   */
  async getFilters(): Promise<models.IFilter[]> {
    try {
      const response = await this.report.service.hpm.get<models.IFilter[]>(`/report/pages/${this.name}/filters`, { uid: this.report.config.uniqueId }, this.report.iframe.contentWindow);
      return response.body;
    } catch (response) {
      throw response.body;
    }
  }

  /**
   * Delete the page from the report
   *
   * ```javascript
   * // Delete the page from the report
   * page.delete();
   * ```
   *
   * @returns {Promise<void>}
   */
  async delete(): Promise<void> {
    try {
      const response = await this.report.service.hpm.delete<void>(`/report/pages/${this.name}`, { }, { uid: this.report.config.uniqueId }, this.report.iframe.contentWindow);
      return response.body;
    } catch (response) {
      throw response.body;
    }
  }

  /**
   * Removes all filters from this page of the report.
   *
   * ```javascript
   * page.removeFilters();
   * ```
   *
   * @returns {Promise<IHttpPostMessageResponse<void>>}
   */
  async removeFilters(): Promise<IHttpPostMessageResponse<void>> {
    return await this.setFilters([]);
  }

  /**
   * Makes the current page the active page of the report.
   *
   * ```javascripot
   * page.setActive();
   * ```
   *
   * @returns {Promise<IHttpPostMessageResponse<void>>}
   */
  async setActive(): Promise<IHttpPostMessageResponse<void>> {
    const page: models.IPage = {
      name: this.name,
      displayName: null,
      isActive: true
    };

    try {
      return await this.report.service.hpm.put<void>('/report/pages/active', page, { uid: this.report.config.uniqueId }, this.report.iframe.contentWindow);
    } catch (response) {
      throw response.body;
    }
  }

  /**
   * Sets all filters on the current page.
   *
   * ```javascript
   * page.setFilters(filters);
   *   .catch(errors => { ... });
   * ```
   *
   * @param {(models.IFilter[])} filters
   * @returns {Promise<IHttpPostMessageResponse<void>>}
   */
  async setFilters(filters: models.IFilter[]): Promise<IHttpPostMessageResponse<void>> {
    try {
      return await this.report.service.hpm.put<void>(`/report/pages/${this.name}/filters`, filters, { uid: this.report.config.uniqueId }, this.report.iframe.contentWindow);
    } catch (response) {
      throw response.body;
    }
  }

  /**
   * Set displayName to the current page.
   *
   * ```javascript
   * page.setName(displayName);
   * ```
   *
   * @returns {Promise<IHttpPostMessageResponse<void>>}
   */
  async setDisplayName(displayName: string): Promise<IHttpPostMessageResponse<void>> {
    const page: models.IPage = {
      name: this.name,
      displayName,
    };

    try {
      return await this.report.service.hpm.put<void>(`/report/pages/${this.name}/name`, page, { uid: this.report.config.uniqueId }, this.report.iframe.contentWindow);
    } catch (response) {
      throw response.body;
    }
  }

  /**
   * Gets all the visuals on the page.
   *
   * ```javascript
   * page.getVisuals()
   *   .then(visuals => { ... });
   * ```
   *
   * @returns {Promise<VisualDescriptor[]>}
   */
  async getVisuals(): Promise<VisualDescriptor[]> {
    if (utils.isRDLEmbed(this.report.config.embedUrl)) {
      return Promise.reject(errors.APINotSupportedForRDLError);
    }

    try {
      const response = await this.report.service.hpm.get<models.IVisual[]>(`/report/pages/${this.name}/visuals`, { uid: this.report.config.uniqueId }, this.report.iframe.contentWindow);
      return response.body
          .map(visual => {
            return new VisualDescriptor(this, visual.name, visual.title, visual.type, visual.layout);
          });
      } catch (response) {
        throw response.body;
      }
  }

  /**
   * Checks if page has layout.
   *
   * ```javascript
   * page.hasLayout(layoutType)
   *  .then(hasLayout: boolean => { ... });
   * ```
   *
   * @returns {(Promise<boolean>)}
   */
  async hasLayout(layoutType): Promise<boolean> {
    if (utils.isRDLEmbed(this.report.config.embedUrl)) {
      return Promise.reject(errors.APINotSupportedForRDLError);
    }

    let layoutTypeEnum = models.LayoutType[layoutType];
    try {
      const response = await this.report.service.hpm.get<boolean>(`/report/pages/${this.name}/layoutTypes/${layoutTypeEnum}`, { uid: this.report.config.uniqueId }, this.report.iframe.contentWindow);
      return response.body;
    } catch (response) {
      throw response.body;
    }
  }
}