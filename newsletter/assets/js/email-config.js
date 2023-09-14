/**
 * @license Copyright (c) 2003-2023, CKSource Holding sp. z o.o. All rights reserved.
 * For licensing, see https://ckeditor.com/legal/ckeditor-oss-license
 */

CKEDITOR.editorConfig = function( config ) {
	// Define changes to default configuration here.
	// For complete reference see:
	// https://ckeditor.com/docs/ckeditor4/latest/api/CKEDITOR_config.html

	// The toolbar groups arrangement, optimized for two toolbar rows.
	config.toolbarGroups = [
		{ name: 'clipboard',   groups: [ 'clipboard', 'undo' ] },
		{ name: 'links' },
		{ name: 'document',	   groups: [ 'mode', 'document', 'doctools' ] },
		'/',
		{ name: 'basicstyles', groups: [ 'basicstyles'] },
		{ name: 'paragraph', groups: ['indent', 'align'] },
		'/',
		{ name: 'styles' },
		{ name: 'colors' },
		{ name: 'about' }
	];

	config.removePlugins = 'docprops,exportpdf';
	config.removeButtons = 'Templates,Save,Print,Flash,NewPage,PasteFromWord,Strike,Anchor,Subscript,Superscript';
	config.enterMode = CKEDITOR.ENTER_BR;
	config.format_tags = 'p;h1;h2;h3;pre';

	// Simplify the dialog windows.
	// config.removeDialogTabs = 'image:advanced;link:advanced';
};
