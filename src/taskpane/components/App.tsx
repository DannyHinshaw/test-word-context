import { Button, ButtonType } from "office-ui-fabric-react";
import * as React from "react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";

/* global Button Header, HeroList, HeroListItem, Progress, Word */

export interface AppProps {
	title: string;
	isOfficeInitialized: boolean;
}

export interface AppState {
	listItems: HeroListItem[];
}

export default class App extends React.Component<AppProps, AppState> {
	constructor(props, context) {
		super(props, context);
		this.state = {
			listItems: []
		};
	}

	componentDidMount() {
		OfficeExtension.config.extendedErrorLogging = true;
		this.setState({
			listItems: [
				{
					icon: "Ribbon",
					primaryText: "Achieve more with Office integration"
				},
				{
					icon: "Unlock",
					primaryText: "Unlock features and functionality"
				},
				{
					icon: "Design",
					primaryText: "Create and visualize like a pro"
				}
			]
		});
	}

	click = async () => {
		// Run a batch operation against the Word object model.
		Word.run(function(context) {

			// Queue a command to get the current selection and then
			// create a proxy range object with the results.
			var range = context.document.getSelection();

			// Queue a command to get the OOXML of the current selection.
			var ooxml = range.getOoxml();

			// Synchronize the document state by executing the queued commands,
			// and return a promise to indicate task completion.
			return context.sync().then(function() {
				console.log("The OOXML read from the document was:  " + ooxml.value);
			});
		})
			.catch(function(error) {
				console.log("Error: " + JSON.stringify(error));
				if (error instanceof OfficeExtension.Error) {
					console.log("Debug info: " + JSON.stringify(error.debugInfo));
				}
			});
	};

	render() {
		const { title, isOfficeInitialized } = this.props;

		if (!isOfficeInitialized) {
			return (
				<Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
			);
		}

		return (
			<div className="ms-welcome">
				<Header logo="assets/logo-filled.png" title={this.props.title} message="Welcome" />
				<HeroList message="Discover what Office Add-ins can do for you today!" items={this.state.listItems}>
					<p className="ms-font-l">
						Modify the source files, then click <b>Run</b>.
					</p>
					<Button
						className="ms-welcome__action"
						buttonType={ButtonType.hero}
						iconProps={{ iconName: "ChevronRight" }}
						onClick={this.click}
					>
						Run
					</Button>
				</HeroList>
			</div>
		);
	}
}
