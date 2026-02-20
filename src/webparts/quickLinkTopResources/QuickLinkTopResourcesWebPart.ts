import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneField,
  PropertyPaneFieldType
} from '@microsoft/sp-property-pane';
import styles from './QuickLinkTopResourcesWebPart.module.scss';

// Interface for a single custom link
interface ICustomLink {
  id: string;
  iconUrl: string;
  title: string;
  url: string;
}

export interface IQuickLinksTopResourcesWebPartProps {
  customLinks: ICustomLink[];
}

interface ILinkItem {
  Title: string;
  LinkUrl: { Url: string };
  Category: string;
}

export default class QuickLinksTopResourcesWebPart extends BaseClientSideWebPart<IQuickLinksTopResourcesWebPartProps> {

  protected onInit(): Promise<void> {
    // Initialize customLinks array if not exists
    if (!this.properties.customLinks) {
      this.properties.customLinks = [];
    }
    return Promise.resolve();
  }

  public render(): void {
    const customLinks = this.properties.customLinks || [];
    
    // Build custom links HTML
    const customLinksHtml = customLinks
      .filter(link => link.title && link.url)
      .map(link => `
        <a href="${link.url}" target="_blank" class="${styles.card}">
          ${link.iconUrl 
            ? `<img src="${link.iconUrl}" alt="icon" class="${styles.customIcon}" />`
            : `<div class="${styles.icon}">★</div>`
          }
          <div class="${styles.title}">${link.title}</div>
        </a>
      `).join('');

    this.domElement.innerHTML = `
      <div class="${styles.webpartContainer}">
        <div class="${styles.section}">
          <div class="${styles.sectionHeader}">
            <div class="${styles.headerAccent}"></div>
            <h1>TOP RESOURCES</h1>
          </div>
          <div id="topResources" class="${styles.grid}">${customLinksHtml}<span id="topResourcesItems">Loading...</span></div>
        </div>
      </div>
    `;

    this.loadLinks();
  }

  private loadLinks(): void {
    const url =
      `${this.context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists/getbytitle('Intranet Links')/items?$select=Title,LinkUrl,Category`;

    this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then(res => {
        if (!res.ok) {
          throw new Error(`Request failed: ${res.status} ${res.statusText}`);
        }
        return res.json();
      })
      .then(data => {
        const items = Array.isArray(data?.value) ? (data.value as ILinkItem[]) : [];
        const topResources = items.filter((i: ILinkItem) => i.Category === 'TOP RESOURCES');
        this.renderCards('topResourcesItems', topResources);
      })
      .catch(() => {
        this.renderCards('topResourcesItems', []);
      });
  }

  private renderCards(containerId: string, items: ILinkItem[]): void {
    const container = this.domElement.querySelector(`#${containerId}`) as HTMLElement | null;

    if (!container) {
      return;
    }

    if (!items || items.length === 0) {
      container.innerHTML = '';
      return;
    }

    container.innerHTML = items.map(item => `
      <a href="${item.LinkUrl.Url}" target="_blank" class="${styles.card}">
        <div class="${styles.icon}">★</div>
        <div class="${styles.title}">${item.Title}</div>
      </a>
    `).join('');
  }

  // Generate unique ID
  private _generateId(): string {
    return 'link_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
  }

  // Add new link entry
  private _handleAddLink(): void {
    if (!this.properties.customLinks) {
      this.properties.customLinks = [];
    }
    
    this.properties.customLinks.push({
      id: this._generateId(),
      iconUrl: '',
      title: '',
      url: ''
    });
    
    this.context.propertyPane.refresh();
    this.render();
  }

  // Handle file upload for specific link
  private _handleFileUpload(linkId: string, file: File): void {
    const reader = new FileReader();
    reader.onload = (e: ProgressEvent<FileReader>) => {
      const result = e.target?.result as string;
      const link = this.properties.customLinks.find(l => l.id === linkId);
      if (link) {
        link.iconUrl = result;
        this.context.propertyPane.refresh();
        this.render();
      }
    };
    reader.readAsDataURL(file);
  }

  // Reset specific link
  private _handleResetLink(linkId: string): void {
    const link = this.properties.customLinks.find(l => l.id === linkId);
    if (link) {
      link.iconUrl = '';
      link.title = '';
      link.url = '';
      this.context.propertyPane.refresh();
      this.render();
    }
  }

  // Delete specific link
  private _handleDeleteLink(linkId: string): void {
    const index = this.properties.customLinks.findIndex(l => l.id === linkId);
    if (index > -1) {
      this.properties.customLinks.splice(index, 1);
      this.context.propertyPane.refresh();
      this.render();
    }
  }

  // Create Add button field
  private _createAddButtonField(): IPropertyPaneField<{}> {
    return {
      type: PropertyPaneFieldType.Custom,
      targetProperty: 'addButton',
      properties: {
        key: 'addButton',
        onRender: this._renderAddButton.bind(this)
      }
    };
  }

  // Render Add button
  private _renderAddButton(elem: HTMLElement): void {
    elem.innerHTML = `
      <button id="addNewLinkBtn" style="
        width: 100%;
        padding: 12px 16px;
        background: #f2b705;
        color: #000;
        border: none;
        border-radius: 4px;
        cursor: pointer;
        font-size: 14px;
        font-weight: 600;
        margin-top: 20px;
      ">Add New Link</button>
    `;

    const addBtn = elem.querySelector('#addNewLinkBtn') as HTMLButtonElement;
    addBtn.addEventListener('click', () => {
      this._handleAddLink();
    });
  }

  // Create link entry field
  private _createLinkEntryField(link: ICustomLink, index: number): IPropertyPaneField<{}> {
    return {
      type: PropertyPaneFieldType.Custom,
      targetProperty: `linkEntry_${link.id}`,
      properties: {
        key: `linkEntry_${link.id}`,
        onRender: (elem: HTMLElement) => this._renderLinkEntry(elem, link, index)
      }
    };
  }

  // Render individual link entry
  private _renderLinkEntry(elem: HTMLElement, link: ICustomLink, index: number): void {
    elem.innerHTML = `
      <div style="border: 1px solid #ddd; border-radius: 8px; padding: 16px; margin-bottom: 16px; background: #f9f9f9;">
        <div style="font-weight: 600; margin-bottom: 12px; color: #333;">Link ${index + 1}</div>
        
        <!-- Buttons row -->
        <div style="display: flex; gap: 8px; margin-bottom: 12px;">
          <button class="uploadBtn" data-id="${link.id}" style="
            flex: 1;
            padding: 8px 12px;
            background: #ffffff;
            color: #000;
            border: 2px solid #f2b705;
            border-radius: 4px;
            cursor: pointer;
            font-size: 12px;
          ">Upload</button>
          <button class="resetBtn" data-id="${link.id}" style="
            flex: 1;
            padding: 8px 12px;
            background: #ffffff;
            color: #000;
            border: 2px solid #f2b705;
            border-radius: 4px;
            cursor: pointer;
            font-size: 12px;
          ">Reset</button>
          <button class="deleteBtn" data-id="${link.id}" style="
            flex: 1;
            padding: 8px 12px;
            background: #ffffff;
            color: #000;
            border: 2px solid #f2b705;
            border-radius: 4px;
            cursor: pointer;
            font-size: 12px;
          ">Delete</button>
        </div>
        
        <input type="file" class="fileInput" data-id="${link.id}" accept="image/*" style="display: none;" />
        
        ${link.iconUrl ? `
          <div style="margin-bottom: 12px; text-align: center;">
            <img src="${link.iconUrl}" alt="Icon Preview" style="max-width: 80px; max-height: 80px; border-radius: 8px; border: 1px solid #ddd;" />
          </div>
        ` : ''}
        
        <!-- Title input -->
        <div style="margin-bottom: 12px;">
          <label style="display: block; font-size: 12px; color: #666; margin-bottom: 4px;">Title</label>
          <input type="text" class="titleInput" data-id="${link.id}" value="${link.title || ''}" placeholder="Enter link title..." style="
            width: 100%;
            padding: 8px 12px;
            border: 1px solid #ddd;
            border-radius: 4px;
            font-size: 14px;
            box-sizing: border-box;
          " />
        </div>
        
        <!-- URL input -->
        <div>
          <label style="display: block; font-size: 12px; color: #666; margin-bottom: 4px;">URL</label>
          <input type="text" class="urlInput" data-id="${link.id}" value="${link.url || ''}" placeholder="https://example.com" style="
            width: 100%;
            padding: 8px 12px;
            border: 1px solid #ddd;
            border-radius: 4px;
            font-size: 14px;
            box-sizing: border-box;
          " />
        </div>
      </div>
    `;

    // Attach event listeners
    const uploadBtn = elem.querySelector(`.uploadBtn[data-id="${link.id}"]`) as HTMLButtonElement;
    const resetBtn = elem.querySelector(`.resetBtn[data-id="${link.id}"]`) as HTMLButtonElement;
    const deleteBtn = elem.querySelector(`.deleteBtn[data-id="${link.id}"]`) as HTMLButtonElement;
    const fileInput = elem.querySelector(`.fileInput[data-id="${link.id}"]`) as HTMLInputElement;
    const titleInput = elem.querySelector(`.titleInput[data-id="${link.id}"]`) as HTMLInputElement;
    const urlInput = elem.querySelector(`.urlInput[data-id="${link.id}"]`) as HTMLInputElement;

    uploadBtn.addEventListener('click', () => {
      fileInput.click();
    });

    fileInput.addEventListener('change', (e: Event) => {
      const target = e.target as HTMLInputElement;
      if (target.files && target.files[0]) {
        this._handleFileUpload(link.id, target.files[0]);
      }
    });

    resetBtn.addEventListener('click', () => {
      this._handleResetLink(link.id);
    });

    deleteBtn.addEventListener('click', () => {
      this._handleDeleteLink(link.id);
    });

    titleInput.addEventListener('input', (e: Event) => {
      const target = e.target as HTMLInputElement;
      const linkItem = this.properties.customLinks.find(l => l.id === link.id);
      if (linkItem) {
        linkItem.title = target.value;
        this.render();
      }
    });

    urlInput.addEventListener('input', (e: Event) => {
      const target = e.target as HTMLInputElement;
      const linkItem = this.properties.customLinks.find(l => l.id === link.id);
      if (linkItem) {
        linkItem.url = target.value;
        this.render();
      }
    });
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const customLinks = this.properties.customLinks || [];
    
    // Build fields array: Link entries first, then Add button at the bottom
    const groupFields: IPropertyPaneField<{}>[] = [
      ...customLinks.map((link, index) => this._createLinkEntryField(link, index)),
      this._createAddButtonField()
    ];

    return {
      pages: [
        {
          header: {
            description: 'Configure Top Resources Links'
          },
          groups: [
            {
              groupName: 'Manage Links',
              groupFields: groupFields
            }
          ]
        }
      ]
    };
  }
}
