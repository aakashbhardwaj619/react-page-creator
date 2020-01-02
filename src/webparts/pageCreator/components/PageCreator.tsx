import * as React from 'react';
import styles from './PageCreator.module.scss';
import { IPageCreatorProps } from './IPageCreatorProps';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { IPageCreatorState } from './IPageCreatorState';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { Dropdown } from 'office-ui-fabric-react/lib/Dropdown';
import { SPService } from '../../../service/SPService';

export default class PageCreator extends React.Component<IPageCreatorProps, IPageCreatorState> {
	constructor(props: IPageCreatorProps) {
		super(props);

		this.state = {
			showPanel: false,
			featuredSiteProperties: [],
			followedSiteProperties: [],
			followedSitesLinkText: 'Show More',
			featuredSitesLinkText: 'Show More',
			showTemplates: false,
			loading: false,
			templateOptions: [],
			selectedSiteUrl: '',
			selectedTemplateId: null,
			pageType: 'SitePage'
		};
	}

	public componentDidMount() {
		this.updateSiteProperties();
		this.updateFollowedSiteProperties();
	}

	public componentDidUpdate(prevProps: IPageCreatorProps) {
		if ((prevProps.selectedSites !== this.props.selectedSites)) {
			this.updateSiteProperties();
		}
		if (prevProps.followedSites !== this.props.followedSites) {
			this.updateFollowedSiteProperties();
		}
	}

	private updateSiteProperties() {
		let siteProperties: any[] = [];
		if (this.props.selectedSites !== undefined) {
			this.props.selectedSites.map(async (site) => {
				let index = site.indexOf('###');
				let siteUrl = site.substring(0, index);
				let siteTitle = site.substring(index + 3);
				let siteLogoUrl = await SPService.GETSITELOGO(this.props.context.spHttpClient, siteUrl);
				siteProperties.push({ siteUrl, siteTitle, siteLogoUrl });
			});
		}
		this.setState({ featuredSiteProperties: siteProperties });
	}

	private updateFollowedSiteProperties() {
		let followedSiteProperties: any[] = [];
		if (this.props.followedSites !== undefined) {
			this.props.followedSites.map(async (site) => {
				let index = site.key.toString().indexOf('###');
				let siteUrl = site.key.toString().substring(0, index);
				let siteTitle = site.key.toString().substring(index + 3);
				let siteLogoUrl = await SPService.GETSITELOGO(this.props.context.spHttpClient, siteUrl);
				followedSiteProperties.push({ siteUrl, siteTitle, siteLogoUrl });
			});
		}
		this.setState({ followedSiteProperties });
	}

  /**
   * Get templates for the selected site
   *
   * @param siteUrl Site URL for which templates need to be retrieved
   */
	private async getTemplates(siteUrl: string) {
		let actualSiteUrl: string = siteUrl;//.substring(0, siteUrl.indexOf('/_layouts'));
		let defaultImageUrl: string = `${actualSiteUrl}/_layouts/15/images/sitepagethumbnail.png`;
		this.setState({ showTemplates: true, selectedSiteUrl: actualSiteUrl });
		let templateOptions: IChoiceGroupOption[] = [];

		templateOptions.push({ key: null, text: 'Blank', imageSrc: defaultImageUrl, selectedImageSrc: defaultImageUrl, imageSize: { width: 120, height: 80 } });

		let siteTemplates = await SPService.GETPAGETEMPLATES(this.props.context.spHttpClient, actualSiteUrl);

		siteTemplates.map((template) => {
			templateOptions.push({ key: template.Id, text: template.Title, imageSrc: template.BannerImageUrl ? template.BannerImageUrl : defaultImageUrl, selectedImageSrc: template.BannerImageUrl ? template.BannerImageUrl : defaultImageUrl, imageSize: { width: 120, height: 80 } });
		});
		this.setState({ templateOptions });
	}

  /**
   * Render Site elements for specified sites
   *
   * @param siteProps Properties for the site elements to be displayed
   */
	private renderSites(siteProps: any[]) {
		return (
			<div>
				{siteProps.map((site) => {
					return (
						<div className={styles.siteTitle}>
							{site.siteLogoUrl &&
								<div className={styles.siteLogo} style={{
									backgroundImage: `url("${site.siteLogoUrl}")`
								}} />
							}
							{!site.siteLogoUrl &&
								<div className={styles.siteLogoText}>
									{site.siteTitle[0]}
								</div>
							}
							<div className={styles.siteAnchor} onClick={() => this.getTemplates(site.siteUrl)}>
								<span className={styles.siteAnchorTitle}>{site.siteTitle}</span>
							</div>
							<br />
						</div>
					);
				})
				}
			</div>
		);
	}

	private followedSiteLinkClicked = (): void => {
		this.setState({
			followedSitesLinkText: this.state.followedSitesLinkText === 'Show More' ? 'Show Less' : 'Show More'
		});
	}

	private featuredSiteLinkClicked = (): void => {
		this.setState({
			featuredSitesLinkText: this.state.featuredSitesLinkText === 'Show More' ? 'Show Less' : 'Show More'
		});
	}

  /**
   * Create News Post
   */
	private createNewPage = async (): Promise<void> => {
		this.setState({ loading: true });
		let newPageUrl: string;
		if (this.state.selectedTemplateId === null) {
			newPageUrl = await SPService.ADDNEWPAGE(this.props.context.spHttpClient, this.state.selectedSiteUrl, this.state.pageType);
		} else {
			newPageUrl = await SPService.ADDNEWPAGEFROMTEMPLATE(this.props.context.spHttpClient, this.state.selectedSiteUrl, this.state.selectedTemplateId, this.state.pageType);
		}
		window.open(`${newPageUrl}?Mode=Edit`, '_blank');
		this.resetState();
	}

	/**
	 * Reset the panel state variables to initial values
	 */
	private resetState = (): void => {
		this.setState({
			showPanel: false,
			selectedSiteUrl: '',
			selectedTemplateId: null,
			showTemplates: false,
			loading: false,
			pageType: 'SitePage',
			templateOptions: [],
			followedSitesLinkText: 'Show More',
			featuredSitesLinkText: 'Show More'
		});
	}

	public render(): React.ReactElement<IPageCreatorProps> {
		return (
			<div className={styles.pageCreator}>
				<div className={styles.container}>
					<div className={styles.row}>
						<div className={styles.column}>
							<PrimaryButton text={this.props.buttonText}
								onClick={() => { this.setState({ showPanel: true }); }}
								styles={{ root: { padding: '10px', float: this.props.buttonAlignment } }}
							/>
							<Panel headerText={this.props.panelHeading}
								isOpen={this.state.showPanel}
								type={PanelType.medium}
								onDismiss={this.resetState}
							>
								{!this.state.showTemplates &&
									<div className={styles.pageCreator}>
										<br />
										{this.state.featuredSiteProperties.length > 0 &&
											<div>
												<span className={styles.categoryTitle}>{this.props.featuredSitesHeading}</span><br /><br />
												{this.state.featuredSitesLinkText === 'Show More' && this.renderSites(this.state.featuredSiteProperties.slice(0, 5))}
												{this.state.featuredSitesLinkText === 'Show Less' && this.renderSites(this.state.featuredSiteProperties)}
												{this.state.featuredSiteProperties.length > 5 &&
													<Link onClick={this.featuredSiteLinkClicked}>{this.state.featuredSitesLinkText}</Link>
												}
												<br /><br />
											</div>
										}
										{this.state.followedSiteProperties.length > 0 &&
											<div>
												<span className={styles.categoryTitle}>Followed Sites</span><br /><br />
												{this.state.followedSitesLinkText === 'Show More' && this.renderSites(this.state.followedSiteProperties.slice(0, 5))}
												{this.state.followedSitesLinkText === 'Show Less' && this.renderSites(this.state.followedSiteProperties)}
												{this.state.followedSiteProperties.length > 5 &&
													<Link onClick={this.followedSiteLinkClicked}>{this.state.followedSitesLinkText}</Link>
												}
											</div>
										}
									</div>
								}
								{this.state.showTemplates &&
									<div>
										<Dropdown label='Select page type'
											defaultSelectedKey='SitePage'
											options={[
												{ key: 'SitePage', text: 'Site Page' },
												{ key: 'NewsPost', text: 'News Post' }
											]}
											onChange={(ev, val) => { this.setState({ pageType: val.key.toString() }); }}
										/>
										<br />
										<ChoiceGroup options={this.state.templateOptions}
											label='Choose Message Template'
											onChange={(ev, val) => { this.setState({ selectedTemplateId: val.key }); }}
										/>
										<br />
										{this.state.loading === true &&
											<Spinner label='Creating new message' ariaLive='assertive' />
										}
										<br />
										<PrimaryButton onClick={this.createNewPage} style={{ marginRight: '8px' }}>
											Create
                    </PrimaryButton>
										<DefaultButton onClick={() => this.setState({ showTemplates: false })}>
											Back
                    </DefaultButton>
									</div>
								}
							</Panel>
						</div>
					</div>
				</div>
			</div>
		);
	}
}
