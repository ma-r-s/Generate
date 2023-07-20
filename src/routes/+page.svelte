<script>
	import Help from '~icons/material-symbols/help-outline-rounded';
	import Excel from '~icons/file-icons/microsoft-excel';
	import Word from '~icons/file-icons/microsoft-word';
	import Down from '~icons/material-symbols/download-rounded';
	import Up from '~icons/material-symbols/upload-rounded';
	import * as XLSX from 'xlsx';
	import Docxtemplater from 'docxtemplater';
	import PizZip from 'pizzip';
	import FileSaver from 'file-saver';
	import DocxMerger from 'docx-merger';
	//Program variables
	let table;
	let document;
	let json;
	let menu = 0;
	let date;
	let fin = [];

	let process = async () => {
		// Load the Excel file
		let fileBuffer = await table[0].arrayBuffer();
		const workbook = XLSX.read(fileBuffer);
		const sheetName = workbook.SheetNames[0];
		const sheet = workbook.Sheets[sheetName];

		// Get the range of cells
		json = XLSX.utils.sheet_to_json(sheet, {
			raw: false
		});

		menu += 1;
	};

	let generate = async () => {
		// Create a new document to store the populated pages

		for (let i = 1; i < json.length; i++) {
			json[i]['actualDate'] = Object.keys(json[i])[date];
			const actual = Object.values(json[i])[date];
			json[i]['actualReading'] = actual;

			json[i]['pastDate'] = Object.keys(json[0])[date - 1];
			const past = Object.values(json[i])[date - 1];
			json[i]['pastReading'] = past;

			const consumed = actual - past;
			json[i]['consumed'] = consumed;
			const price = Object.values(json[0])[date];
			json[i]['ppu'] = price.toLocaleString('en-US', { style: 'currency', currency: 'USD' });
			const value = Math.ceil(consumed * price);
			json[i]['value'] = value.toLocaleString('en-US', { style: 'currency', currency: 'USD' });

			// Clone the template for each page

			let fileBuffer = await document[0].arrayBuffer();
			const zip = new PizZip(fileBuffer);
			const template = new Docxtemplater();
			template.loadZip(zip);
			template.setData(json[i]);
			template.render();
			// Get the generated page content
			fin.push(template.getZip().generate({ type: 'uint8array' }));
		}
		// Generate the final document content
		let docx = new DocxMerger({}, fin);
		// Write the code here
		docx.save('blob', function (data) {
			console.log(Object.keys(json[0])[date]);
			FileSaver.saveAs(data, 'Facturas-' + Object.keys(json[0])[date] + '.docx');
		});
		// Save the Blob as a file using FileSaver
	};
</script>

<button class="top-2 right-2 absolute" onclick="my_modal_1.showModal()">
	<Help class="w-7 h-7" />
</button>
<dialog id="my_modal_1" class="modal">
	<form method="dialog" class="modal-box">
		<button class="btn btn-sm btn-circle btn-ghost absolute right-2 top-2">✕</button>
		<h3 class="font-bold text-lg mb-4">Generador de facturas 3000</h3>
		Si usted se llama Alejandro Ruiz, llame al 3186851696.
		<img class="my-4" src="./Photo.jpg" />
		<a href="https://github.com/ma-r-s/Generate" target="_blank" class="btn btn-primary">Source</a>
	</form>
</dialog>

<div class="flex flex-col items-center justify-center h-screen gap-9">
	<h1 class="font-bold text-3xl pb-8">Generador de facturas</h1>
	{#if menu == 0}
		<label class="btn w-80">
			<input type="file" class="hidden" bind:files={document} />

			<Word class="w-7 h-7 mr-3" />
			{#if document && document[0]}
				{document[0].name.slice(0, 20)}
				{#if document[0].name.length > 20}
					...
				{/if}
			{:else}
				Seleccionar plantilla
			{/if}
		</label>

		<label class="btn w-80">
			<input type="file" class="hidden" bind:files={table} />

			<Excel class="w-7 h-7 mr-3" />
			{#if table && table[0]}
				{table[0].name.slice(0, 20)}
				{#if table[0].name.length > 20}
					...
				{/if}
			{:else}
				Seleccionar tabla
			{/if}
		</label>
		<button on:click={process} class="btn btn-primary {table && document ? '' : 'btn-disabled'}">
			Procesar archivos
			<Up class="w-8 h-8 ml-2" />
		</button>
	{:else}
		<select class="select w-full max-w-xs select-bordered" bind:value={date}>
			<option disabled selected>Elegir fecha</option>
			{#each Object.keys(json[0]).slice(2) as key, i}
				<option value={i + 2}>{key}</option>
			{/each}
		</select>
		<div class="form-control w-72">
			<label class="label cursor-pointer">
				<span class="label-text">Incluir mora</span>
				<input type="checkbox" checked="checked" class="checkbox" />
			</label>
		</div>
		<button on:click={generate} class="btn btn-accent"
			>Generar
			<Down class="w-8 h-8 ml-2" />
		</button>
	{/if}
</div>
