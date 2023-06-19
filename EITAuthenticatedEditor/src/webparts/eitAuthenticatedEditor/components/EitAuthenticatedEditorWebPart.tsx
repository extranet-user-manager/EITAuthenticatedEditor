import * as React from 'react';
import { Helmet } from "react-helmet";
import styles from './EitAuthenticatedEditorWebPart.module.scss';
import * as strings from 'EitAuthenticatedEditorWebPartStrings';
import { IEitAuthenticatedEditorWebPartProps } from './IEitAuthenticatedEditorWebPartProps';
import { IEitAuthenticatedEditorWebPartState } from './IEitAuthenticatedEditorWebPartState';
import { ColorClassNames } from '@uifabric/styling';
import { Guid } from '@microsoft/sp-core-library';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { Stack } from 'office-ui-fabric-react/lib/Stack';
import { Text } from 'office-ui-fabric-react/lib/Text';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/components/MessageBar';
import { Spinner } from 'office-ui-fabric-react/lib/components/Spinner';
import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';
import InnerHTML from 'dangerously-set-html-content';

export default class EitAuthenticatedEditorWebPart extends React.Component<IEitAuthenticatedEditorWebPartProps, IEitAuthenticatedEditorWebPartState> {

    private TemplateRendered: boolean = false;
    private ResourcesRendered: boolean = false;

    constructor(props: IEitAuthenticatedEditorWebPartProps) {
        super(props);

        this.state = {
            resourcesRendered: false
        };
    }

    public render(): React.ReactElement<IEitAuthenticatedEditorWebPartProps> {
        return (
            <Fabric className={styles.eitAuthenticatedEditorWebPart}>
                {(this.props.WebPartTitle) ? this.RenderWebPartTitle() : null}
                {this.RenderTemplate()}
            </Fabric>
        );
    }

    private RenderWebPartTitle(): JSX.Element {
        return (
            <span className={styles.webpartTitle}>{this.props.WebPartTitle.replace("{SiteCollectionTitle}", this.props.SiteCollectionTitle)}</span>
        );
    }

    private RenderSpinner(): JSX.Element {
        let text = strings.LoadingText;
        return (
            <Spinner label={text} />
        );
    }

    private RenderErrors(): JSX.Element {
        return (
            <MessageBar messageBarType={MessageBarType.error} isMultiline={true}>{strings.ErrorText}</MessageBar>
        );
    }

    private RenderSuccess(): JSX.Element {
        return (
            <MessageBar messageBarType={MessageBarType.success} isMultiline={true}>{strings.SuccessText}</MessageBar>
        );
    }

    private RenderScripts(): JSX.Element {
        var resourceArray = this.props.Resources.split(/\r?\n/);
        let elements: any = [];
        if (resourceArray.length) {
            resourceArray.forEach(resource => {
                elements.push(<script src={resource}></script>);
            });
            //this.setState({ resourcesRendered: true });
            this.ResourcesRendered = true;
            return (
                <Helmet>
                    <meta charSet="utf-8" />
                    <title>My Title</title>
                    {elements}
                </Helmet>
            );
        } else {
            return null;
        }
    }

    private myScripts(): JSX.Element {
        return (
            <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
        );
    }

    private RenderTemplate(): JSX.Element {

        console.log(this.props.AadTokenProvider);

        // this.props.AadTokenProvider.getToken(this.props.EumV6ApiAudience).then((accessToken: string): void => {
        //     this.props.HttpClient.get(`${this.props.EumV6ApiUrl}/Users/GetAccessLevel`, HttpClient.configurations.v1,
        //       {
        //         headers: {
        //           'accept': 'application/json',
        //           'authorization': `Bearer ${accessToken}`
        //         }
        //       })

        let Html = `` as string;

        Html += '<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>';

        // if no template is provided then load the sample
        if (!this.props.TemplateUrl) {
            Html = `<p>Please provide template URL.</p>`;
        } else {
            Html += this.props.TemplateHtml;
        }

        

        this.TemplateRendered = true;

        console.log(Html);

        return (
            <div>
                <Helmet>
                    <meta charSet="utf-8" />
                    <title>My Title</title>
                    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
                </Helmet>
                <InnerHTML html={Html} />  
            </div>
        );
    }

}