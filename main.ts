import { App, displayTooltip, Editor, MarkdownView, Modal, Notice, Plugin, PluginSettingTab, Setting, moment } from 'obsidian';
import MsgReader from '@kenjiuno/msgreader';
import proxyData from 'mustache-validator';
const Mustache = require('mustache');

const OutlookMeetingNotesDefaultFilenamePattern =
	'{{#helper_dateFormat}}{{apptStartWhole}}|YYYY-MM-DD HH.mm{{/helper_dateFormat}} {{subject}}';

const OutlookMeetingNotesDefaultTemplate = `---
title: {{subject}}
subtitle: meeting notes
date: {{#helper_dateFormat}}{{apptStartWhole}}|L LT{{/helper_dateFormat}}
meeting: 'true'
meeting-location: {{apptLocation}}
meeting-recipients:
{{#recipients}}
  - {{name}}
{{/recipients}}
meeting-invite: {{body}}
---
`;

interface OutlookMeetingNotesSettings {
	notesFolder: string;
	invalidFilenameCharReplacement: string;
	fileNamePattern: string;
	notesTemplate: string;
}

const DEFAULT_SETTINGS: OutlookMeetingNotesSettings = {
	notesFolder: '',
	invalidFilenameCharReplacement: '',
	fileNamePattern: OutlookMeetingNotesDefaultFilenamePattern,
	notesTemplate: OutlookMeetingNotesDefaultTemplate
}

export default class OutlookMeetingNotes extends Plugin {
	settings: OutlookMeetingNotesSettings;

	//TODO: Add functionality to export meeting notes nicely

	async createMeetingNote(msgPromise: Promise<MsgReader>) {
		const msg = await msgPromise;
		try {
			const { vault } = this.app;

			// Get the file data from MsgReader
			const fileData = msg.getFileData();

			// Check if we got a suitable meeting
			if (fileData.dataType != 'msg') {
				throw new TypeError('Outlook Meeting Notes cannot process the file. '
					+ 'MsgReader did not parse the file as valid msg format.');
			} else if (fileData.messageClass != 'IPM.Appointment') {
				throw new TypeError('Outlook Meeting Notes cannot process the file. '
					+ 'It is a valid msg file but not an appointment or meeting.');
			}

			this.addHelperFunctions(fileData);

			let folderPath = this.settings.notesFolder;
			if (folderPath == '') { folderPath = '/'; }
			const fileNameEscape = {
				escape: (str: string): string => {
					return str.replaceAll('/', this.settings.invalidFilenameCharReplacement);
				}
			}
			const fileNameMustache = Mustache.render(
				this.settings.fileNamePattern,
				proxyData(fileData),
				undefined,
				fileNameEscape)
				.replaceAll(/[*"\\<>:|?]/g, this.settings.invalidFilenameCharReplacement);
			const filePath = folderPath + '/' + fileNameMustache + '.md';
			const newFolderPath = filePath.replace(/\/[^/]*$/, '');
			let meetingNoteFile = vault.getFileByPath(filePath);
			if (meetingNoteFile) {
				// File already exists
				// TODO: send a message to the user
			}
			else {
				if (vault.getFolderByPath(newFolderPath) == null) {
					vault.createFolder(newFolderPath);
				}
				const mustacheOutput = this.renderTemplate(
					this.settings.notesTemplate,
					fileData);
				meetingNoteFile = await vault.create(filePath, mustacheOutput);
				new Notice('New file: ' + meetingNoteFile.basename);
			}
			const openInNewTab = false;
			this.app.workspace.getLeaf(openInNewTab).openFile(meetingNoteFile);
			const fe = this.app.internalPlugins.getEnabledPluginById("file-explorer");
			if (fe) { fe.revealInFolder(meetingNoteFile); }
		} catch (ee: unknown) {
			// TODO: Handle errors reasonably -- differently between msg missing elements and errors creating file
			if (ee instanceof Error) { new Notice(ee.name + ':\n' + ee.message); }
			throw ee;
		}
	}

	// Custom funciton to handle a file being dropped onto the ribbon icon.
	async handleDropEvent(dropevt: DragEvent) {
		if (dropevt.dataTransfer == null) {
			throw new ReferenceError('Outlook Meeting Notes cannot handle the DragEvent. The event had a null '
				+ 'dataTransfer property, which should never happen when dispatched by the browser, according '
				+ 'to https://developer.mozilla.org/en-US/docs/Web/API/DragEvent/dataTransfer');
		} else {
			const droppedFiles = dropevt.dataTransfer.files
			if (droppedFiles.length != 1) {
				new Notice('Outlook Meeting Notes can only handle one meeting being dropped onto the ribbon icon');
			}
			else {
				// One file was dropped, hand it over to createMeetingNote
				const droppedFile = droppedFiles[0];
				new Notice('Dropped: ' + droppedFile.name);

				const fr = new FileReader();
				fr.onload = () => {
					if (fr.result == null) {
						throw new ReferenceError('Outlook Meeting Notes cannot handle the DragEvent. The FileReader had '
							+ 'a null result property, which should not be possible.');
					} else if (!(fr.result instanceof ArrayBuffer)) {
						throw new TypeError('Outlook Meeting Notes cannot handle the DragEvent. The FileReader result '
							+ 'property was not an ArrayBuffer, which should be impossible.');
					} else {
						// As readAsArrayBuffer is being used, below, fr.result will be an ArrayBuffer.
						const msgRdr = new MsgReader(fr.result);
						this.createMeetingNote(msgRdr);
					}
				}
				fr.readAsArrayBuffer(droppedFile)
			}
		}
	}

	async onload() {
		await this.loadSettings();

		const tooltipMessage = 'Outlook Meeting Notes: Drag and drop a meeting onto this icon from Outlook (or a .msg file) to create a meeting note.';

		// This creates an icon in the left ribbon.
		// Create an icon that does nothing when clicked, as the effect is from 
		const ribbonIconEl = this.addRibbonIcon('calendar-clock', tooltipMessage, () => { });

		// These respond to something being dragged over the ribbon icon and dropped onto it
		this.registerEvent(ribbonIconEl.on('dragenter', 'div', () => {
			// Style for icon-dragover is defined in the CSS file.
			ribbonIconEl.toggleClass('is-being-dragged-over', true);
			// Display a tooltip.
			const ttPosition = ribbonIconEl.getAttribute('data-tooltip-position');
			const ttDelay = ribbonIconEl.getAttribute('data-tooltip-delay');
			if (ttPosition != null && ttDelay != null) {
				displayTooltip(ribbonIconEl, tooltipMessage, { placement: ttPosition, delay: Number(ttDelay) });
			} else {
				displayTooltip(ribbonIconEl, tooltipMessage);
			}
		}));
		this.registerEvent(ribbonIconEl.on('dragleave', 'div', () => {
			ribbonIconEl.toggleClass('is-being-dragged-over', false);
			// Remove the tooltip when the user leaves the ribbon icon while still dragging.
			const tooltip = document.getElementsByClassName('tooltip')[0]
			if (tooltip) { tooltip.remove(); }
		}));
		this.registerEvent(ribbonIconEl.on('dragover', 'div', (dragevt) => {
			// User is dragging something over the icon

			// Prevent the default behaviour of not allowing the drop event
			dragevt.preventDefault();
			// The dropEffect changes the cursor. Options are 'none', 'move', 'copy', and 'link'.
			if (dragevt.dataTransfer != null) { dragevt.dataTransfer.dropEffect = 'copy'; }
		}));
		this.registerEvent(ribbonIconEl.on('drop', 'div', (dropevt) => {
			// User has dropped something on the icon

			// If the user drops something, dragleave doesn't get called.
			ribbonIconEl.toggleClass('is-being-dragged-over', false);

			// Prevent the default behaviour because we're handling it.
			dropevt.preventDefault();
			// Call the custom function defined above.
			this.handleDropEvent(dropevt);
		}));
		ribbonIconEl.addClass('outlook-meeting-notes-icon');


		// This adds a status bar item to the bottom of the app. Does not work on mobile apps.
		// const statusBarItemEl = this.addStatusBarItem();
		// statusBarItemEl.setText('Status Bar Text');

		//TODO: Add commands to create meeting notes.

		// This adds a simple command that can be triggered anywhere
		this.addCommand({
			id: 'open-sample-modal-simple',
			name: 'Open sample modal (simple)',
			callback: () => {
				new OutlookMeetingNotesModal(this.app).open();
			}
		});
		// This adds an editor command that can perform some operation on the current editor instance
		this.addCommand({
			id: 'sample-editor-command',
			name: 'OutlookMeetingNotes editor command',
			editorCallback: (editor: Editor, view: MarkdownView) => {
				console.log(editor.getSelection());
				editor.replaceSelection('OutlookMeetingNotes Editor Command');
			}
		});
		// This adds a complex command that can check whether the current state of the app allows execution of the command
		this.addCommand({
			id: 'open-sample-modal-complex',
			name: 'Open sample modal (complex)',
			checkCallback: (checking: boolean) => {
				// Conditions to check
				const markdownView = this.app.workspace.getActiveViewOfType(MarkdownView);
				if (markdownView) {
					// If checking is true, we're simply 'checking' if the command can be run.
					// If checking is false, then we want to actually perform the operation.
					if (!checking) {
						new OutlookMeetingNotesModal(this.app).open();
					}

					// This command will only show up in Command Palette when the check function returns true
					return true;
				}
			}
		});

		// This adds a settings tab so the user can configure various aspects of the plugin
		this.addSettingTab(new OutlookMeetingNotesSettingTab(this.app, this));

		// If the plugin hooks up any global DOM events (on parts of the app that doesn't belong to this plugin)
		// Using this function will automatically remove the event listener when this plugin is disabled.
		// this.registerDomEvent(document, 'click', (evt: MouseEvent) => {
		// 	console.log('click', evt);
		// });

		// When registering intervals, this function will automatically clear the interval when the plugin is disabled.
		//this.registerInterval(window.setInterval(() => console.log('setInterval'), 5 * 60 * 1000));
	}

	onunload() {

	}

	async loadSettings() {
		this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
	}

	async saveSettings() {
		await this.saveData(this.settings);
	}

	addHelperFunctions(hash: any): any {
		const helperFunctions = {
			firstWord: () => {
				return function (words: string, render: any) {
					return render(words).replace(/\W.*$/, '');
				}
			},
			dateFormat: () => {
				return function (datetime_format: string, render: any) {
					const parts = datetime_format.split('|')
					return moment(render(parts[0]).trim()).format(parts[1]);
				}
			}
		};

		for (let func in helperFunctions) {
			hash['helper_' + func] = helperFunctions[func];
		}
		// Add helper functions to all objects in arrays so that the helper 
		// functions work inside mustache sections
		for (let property in hash) {
			if (hash[property] instanceof Array) {
				for (let subproperty in hash[property]) {
					if (hash[property][subproperty] instanceof Object) {
						for (let func in helperFunctions) {
							hash[property][subproperty]['helper_' + func] = helperFunctions[func];
						}
					}
				}
			}
		}

		return hash;
	}

	//////////////////////
	// renderTemplate() //
	//////////////////////
	// Parse template into YAML and markdown sections to use different escaping for each
	renderTemplate(template: string, hash: any): string {
		// Regex /^---(\r\n?|\n).*?(\r\n?|\n)---($|\r\n?|\n)/s matches '---' at the start of the string,
		// then a platform-independent new-line, followed by a non-greedy match (because of *?) of any
		// character including newlines (because of /s at end), followed by another --- on its own line
		// with either another platform-independent new-line afterwards or the end of the string.
		const templateYAMLMatch = template.match(/^---(\r\n?|\n).*?(\r\n?|\n)---($|\r\n?|\n)/s);
		// If there was a match, then make notesTemplateMD everything after it, otherwise it's the whole thing.
		const templateMD = templateYAMLMatch ? template.substring(templateYAMLMatch[0].length) : template;

		let output = ''

		if (templateYAMLMatch) {
			const mustacheYAMLOptions = {
				escape: (str: string): string => {
					// Escape YAML
					const found = str.match(/\r\n?|\n/);
					if (found) {
						return '|\n' + '  ' + str.replaceAll(/\r\n?|\n/g, '\n  ');
					} else if (str.match(/[:#\[\]\{\},]/)) {
						return '"' + str.replaceAll(/["\\]/g, '\\$&') + '"';
					}
					else return str;
				}
			}

			output = output + Mustache.render(
				templateYAMLMatch[0],
				proxyData(hash),
				undefined,
				mustacheYAMLOptions);
		}

		if (templateMD) {
			const mustacheMDOptions = {
				escape: (str: string): string => {
					return str.replaceAll(/[\\\`\*\_\[\]\{\}\<\>\(\)\#\!\|\^]/g, '\\$&')
						.replaceAll('%%', '\\%\\%')
						.replaceAll('~~', '\\~\\~')
						.replaceAll('==', '\\=\\=');
				}
			}

			output = output + Mustache.render(
				templateMD,
				proxyData(hash),
				undefined,
				mustacheMDOptions);
		}

		return output;
	}

}

class OutlookMeetingNotesModal extends Modal {
	constructor(app: App) {
		super(app);
	}

	onOpen() {
		const { contentEl } = this;
		contentEl.setText('Woah!');
	}

	onClose() {
		const { contentEl } = this;
		contentEl.empty();
	}
}

class OutlookMeetingNotesSettingTab extends PluginSettingTab {
	plugin: OutlookMeetingNotes;

	constructor(app: App, plugin: OutlookMeetingNotes) {
		super(app, plugin);
		this.plugin = plugin;
	}

	display(): void {
		const { containerEl } = this;

		containerEl.empty();

		new Setting(containerEl)
			.setName('Folder location')
			.setDesc('Notes will be created in this folder.')
			.addText(text => text
				.setPlaceholder('Example: folder 1/subfolder 2')
				.setValue(this.plugin.settings.notesFolder)
				.onChange(async (value) => {
					this.plugin.settings.notesFolder = value;
					await this.plugin.saveSettings();
				}));

		new Setting(containerEl)
			.setName('Filename pattern')
			.setDesc('This pattern will be used to name new notes.')
			.addText(text => text
				.setPlaceholder('Default: ' + OutlookMeetingNotesDefaultFilenamePattern)
				.setValue(this.plugin.settings.fileNamePattern)
				.onChange(async (value) => {
					if (value == '') {
						this.plugin.settings.fileNamePattern = OutlookMeetingNotesDefaultFilenamePattern;
					} else {
						this.plugin.settings.fileNamePattern = value;
					}
					await this.plugin.saveSettings();
				}));

		new Setting(containerEl)
			.setName('Invalid character substitute')
			.setDesc('This character (or string) will be used in place of any invalid characters for new note filenames.')
			.addText(text => text
				.setPlaceholder('Example: _')
				.setValue(this.plugin.settings.invalidFilenameCharReplacement)
				.onChange(async (value) => {
					this.plugin.settings.invalidFilenameCharReplacement = value;
					await this.plugin.saveSettings();
				}));

		new Setting(containerEl)
			.setName('Template')
			.setDesc('This template will be used for new notes.')
			.addTextArea(text => text
				.setPlaceholder('Default: ' + OutlookMeetingNotesDefaultFilenamePattern)
				.setValue(this.plugin.settings.notesTemplate)
				.onChange(async (value) => {
					this.plugin.settings.notesTemplate = value;
					await this.plugin.saveSettings();
				}));

		new Setting(containerEl)
			.setDesc((() => {
				const df = document.createDocumentFragment();
				df.appendChild(document.createTextNode(
					'For more information about filename patterns and the syntax for templates, see the '
				));
				const link = document.createElement('a');
				link.href = 'https://www.github.com/davidingerslev/outlook-meeting-notes/#outlook-meeting-notes';
				link.target = '_blank';
				link.rel = 'noopener';
				link.textContent = 'documentation';
				df.appendChild(link);
				df.appendChild(document.createTextNode('.'))
				return df;
			})())

	}
}

/*
 * Notes from the sample plugin README.md

## Releasing new releases

- Update your `manifest.json` with your new version number, such as `1.0.1`, and the minimum Obsidian version required for your latest release.
- Update your `versions.json` file with `"new-plugin-version": "minimum-obsidian-version"` so older versions of Obsidian can download an older version of your plugin that's compatible.
- Create new GitHub release using your new version number as the "Tag version". Use the exact version number, don't include a prefix `v`. See here for an example: https://github.com/obsidianmd/obsidian-sample-plugin/releases
- Upload the files `manifest.json`, `main.js`, `styles.css` as binary attachments. Note: The manifest.json file must be in two places, first the root path of your repository and also in the release.
- Publish the release.

> You can simplify the version bump process by running `npm version patch`, `npm version minor` or `npm version major` after updating `minAppVersion` manually in `manifest.json`.
> The command will bump version in `manifest.json` and `package.json`, and add the entry for the new version to `versions.json`

## Adding your plugin to the community plugin list

- Check the [plugin guidelines](https://docs.obsidian.md/Plugins/Releasing/Plugin+guidelines).
- Publish an initial version.
- Make sure you have a `README.md` file in the root of your repo.
- Make a pull request at https://github.com/obsidianmd/obsidian-releases to add your plugin.
 */
