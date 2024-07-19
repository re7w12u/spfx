import * as React from 'react';
import { ProjectsProps } from './ProjectsProps';
import { ProjectsState } from './ProjectsState';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { DocumentCard, DocumentCardDetails, DocumentCardTitle } from '@fluentui/react';
import { PrimaryButton } from '@fluentui/react/lib/Button';


export class Projects extends React.Component<ProjectsProps, ProjectsState> {


    constructor(props: ProjectsProps, state: ProjectsState) {
        super(props);
        this.state = {
            items: []
        }        
    }

    public getItems(filterVal: string) {
        if (filterVal === "*") {
            this.props.context.spHttpClient
                .get(`${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Projects')/Items?$expand=ProjectManager&$select=*,ProjectManager,ProjectManager/EMail,ProjectManager/Title`,
                    SPHttpClient.configurations.v1
                )
                .then(
                    (response: SPHttpClientResponse): Promise<{ value: any }> => {
                        return response.json();
                    }
                )
                .then(
                    (response: { value: any }) => {
                        var _items: any[] = [];
                        _items = _items.concat(response.value);
                        this.setState({ items: _items });
                    }
                )
        }
        else {
            this.props.context.spHttpClient
                .get(`${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Projects')/Items?$expand=ProjectManager&$select=*,ProjectManager,ProjectManager/EMail,ProjectManager/Title&$filter=Status eq %27${filterVal}%27`,
                    SPHttpClient.configurations.v1
                )
                .then(
                    (response: SPHttpClientResponse): Promise<{ value: any }> => {
                        return response.json();
                    }
                )
                .then(
                    (response: { value: any }) => {
                        var _items: any[] = [];
                        _items = _items.concat(response.value);
                        this.setState({ items: _items });
                    }
                )
        }
    }

    public componentDidMount(): void {
        var getAll = "*";
        this.getItems(getAll);
        console.log("Hello from webpart");
    }

    public progFilter(filterVal: string) {
        switch (filterVal) {
            case "*":
                return this.getItems(filterVal);
            case "in progress":
                return this.getItems(filterVal);
            case "completed":
                return this.getItems(filterVal);
            case "cancelled":
                return this.getItems(filterVal);
            case "not started":
                return this.getItems(filterVal);
        }
    }

    public render(): React.ReactElement<ProjectsProps> {
        var _projDocLink = `${this.props.context.pageContext.web.absoluteUrl}/project%20documents/forms/allitems.aspx?FilterField1=Project&FilterValue1=`;

        var notStarted = "not started";
        var inProg = "in progress";
        var completed = "completed";
        var cancelled = "cancelled";
        var allProj = "*";


        return <div>
            <div>
                <PrimaryButton onClick={() => this.progFilter(allProj)} text='All'></PrimaryButton>
                <PrimaryButton onClick={() => this.progFilter(inProg)} text='In Progress'></PrimaryButton>
                <PrimaryButton onClick={() => this.progFilter(notStarted)} text='Not Started'></PrimaryButton>
                <PrimaryButton onClick={() => this.progFilter(completed)} text='Completed'></PrimaryButton>
                <PrimaryButton onClick={() => this.progFilter(cancelled)} text='Cancelled'></PrimaryButton>

                {this.state.items.map((item: any, key: any) =>
                    <DocumentCard>
                        <a href={_projDocLink + item.Title}><DocumentCardTitle title={item.Title}></DocumentCardTitle></a>                        
                        <DocumentCardDetails>
                        <div><a href={"mailto:"+item.ProjectManager.EMail}>{item.ProjectManager.Title}</a></div>
                        <div>{item.Status}</div>
                        </DocumentCardDetails>
                    </DocumentCard>
                )}
            </div>
        </div>
    }

}