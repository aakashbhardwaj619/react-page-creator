import { SPHttpClient, SPHttpClientResponse, MSGraphClient, MSGraphClientFactory } from '@microsoft/sp-http';
import { IPropertyPaneDropdownOption } from '@microsoft/sp-webpart-base';

export class SPService {
	/**
   * Gets and returns the Site Logo URL for the specified site
   *
	 * @param spHttpClient The spHttpClient for current context
   * @param siteUrl The Site URL for which logo needs to be retrieved
   */
	public static async GETSITELOGO(spHttpClient: SPHttpClient, siteUrl: string): Promise<string> {
		let siteLogoResponse = await spHttpClient.get(
			`${siteUrl}/_api/web?$select=SiteLogoUrl`,
			SPHttpClient.configurations.v1,
			{}
		);
		let siteLogo = await siteLogoResponse.json();
		return siteLogo.SiteLogoUrl;
	}

	/**
	 * Get available page templates in the selected site
	 *
	 * @param spHttpClient The spHttpClient for current context
	 * @param siteUrl The Site URL for which logo needs to be retrieved
	 */
	public static async GETPAGETEMPLATES(spHttpClient: SPHttpClient, siteUrl: string): Promise<any[]> {
		let siteTemplatesResponse = await spHttpClient.get(
			`${siteUrl}/_api/sitepages/pages/templates?asjson=1&$select=Id,Title,BannerImageUrl`,
			SPHttpClient.configurations.v1,
			{}
		);

		if (siteTemplatesResponse.ok) {
			let siteTemplates = await siteTemplatesResponse.json();
			return siteTemplates.value;
		} else {
			return [];
		}
	}

	/**
	 * Create a new blank News post in the selected site and returns the new page URL
	 *
	 * @param spHttpClient The spHttpClient for current context
	 * @param siteUrl The Site URL where new page needs to be created
	 */
	public static async ADDNEWPAGE(spHttpClient: SPHttpClient, siteUrl: string, pageType: string): Promise<string> {
		let newPageUrl: string = '';
		let promotedState = pageType === 'SitePage' ? 0 : 1 ;

		const body: string = JSON.stringify({
			'__metadata': {
				'type': 'SP.Publishing.SitePage'
			},
			'PromotedState': promotedState,
			'PageLayoutType': 'Article'
		});

		let dataResponse: SPHttpClientResponse = await spHttpClient.post(
			`${siteUrl}/_api/sitepages/pages`,
			SPHttpClient.configurations.v1,
			{
				headers: {
					'Accept': 'application/json',
					'Content-type': 'application/json;odata=verbose;charset=utf-8',
					'odata-version': '3.0',
					'IF-MATCH': '*',
					'X-HTTP-Method': 'POST'
				},
				body: body
			}
		);
		if (dataResponse.ok) {
			let dataResponseJson = await dataResponse.json();
			newPageUrl = dataResponseJson.AbsoluteUrl;
		}

		return newPageUrl;
	}

	/**
	 * Create a new News post using the selected template and returns the new page URL
	 *
	 * @param spHttpClient The spHttpClient for current context
	 * @param siteUrl The Site URL where new page needs to be created
	 * @param templateId The template ID from which new page will be created
	 */
	public static async ADDNEWPAGEFROMTEMPLATE(spHttpClient: SPHttpClient, siteUrl: string, templateId: string, pageType: string): Promise<string> {
		let newPageUrl: string = '';
		let copyEndpoint = pageType === 'SitePage' ? 'Copy' : 'CreateNewsCopy' ;

		let dataResponse: SPHttpClientResponse = await spHttpClient.post(
			`${siteUrl}/_api/sitepages/Pages/GetById(${templateId})/${copyEndpoint}`,
			SPHttpClient.configurations.v1,
			{
				headers: {
					'Accept': 'application/json',
					'Content-type': 'application/json;odata=verbose;charset=utf-8',
					'X-HTTP-Method': 'POST'
				}
			}
		);
		if (dataResponse.ok) {
			let dataResponseJson = await dataResponse.json();
			newPageUrl = dataResponseJson.AbsoluteUrl;
		}

		return newPageUrl;
	}

	/**
	 * Gets all sites in the tenant using Graph API
	 *
	 * @param msGraphClientFactory MSGraphClientFactory object for the current context
	 */
	public static async GETALLSITES(msGraphClientFactory: MSGraphClientFactory, searchText: string): Promise<IPropertyPaneDropdownOption[]> {
		let selectedSites: IPropertyPaneDropdownOption[] = [];
		
		let searchQuery: string;
		searchQuery = searchText === '*' ?  `*` : `{${searchText}}`;
		
		let _msGraphClient: MSGraphClient = await msGraphClientFactory.getClient();
		let response = await _msGraphClient.api(`/sites?search=${searchQuery}&$select=displayName,webUrl`).get();
		response.value.map((currentSite) => {
			selectedSites.push({ key: `${currentSite.webUrl}###${currentSite.displayName}`, text: currentSite.displayName });
		});
		
		return selectedSites;
	}

	/**
	 * Gets all sites followed by the current user using Graph API
	 *
	 * @param msGraphClientFactory MSGraphClientFactory object for the current context
	 */
	public static async GETFOLLOWEDSITES(msGraphClientFactory: MSGraphClientFactory): Promise<IPropertyPaneDropdownOption[]> {
		let selectedSites: IPropertyPaneDropdownOption[] = [];
		
		let _msGraphClient: MSGraphClient = await msGraphClientFactory.getClient();
		let response = await _msGraphClient.api(`/me/followedSites?$select=displayName,webUrl`).version('beta').get();
		response.value.map((currentSite) => {
			selectedSites.push({ key: `${currentSite.webUrl}###${currentSite.displayName}`, text: currentSite.displayName });
		});
		
		return selectedSites;
	}
}