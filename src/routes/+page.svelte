<script>
	import FileIconsMicrosoftExcel from '~icons/file-icons/microsoft-excel';
	import FileIconsMicrosoftWord from '~icons/file-icons/microsoft-word';
	import IcRoundSaveAlt from '~icons/ic/round-save-alt';
	let table;
	let document;
</script>

<div class="bg-stone-900 overflow-hidden text-white h-screen">
	<div class="grid divide-x-4 h-screen divide-stone-800 grid-cols-3 divide-dashed">
		<label class="custom-file-upload flex-1 flex-col flex justify-center items-center truncate">
			<input type="file" class="hidden" bind:files={table} />
			<FileIconsMicrosoftExcel class="h-16 w-16" />
			{#if table && table[0]}
				<p class="pt-5">
					{table[0].name.slice(0, 14)}
					{#if table[0].name.length > 14}
						...
					{/if}
				</p>
			{:else}
				<p class="pt-5">Seleccionar</p>
			{/if}
		</label>
		<div class="flex-1 flex-col flex justify-center items-center">
			<label class="custom-file-upload flex-1 flex-col flex justify-center items-center truncate">
				<input type="file" class="hidden" bind:files={document} />
				<FileIconsMicrosoftWord class="h-16 w-16" />
				{#if document && document[0]}
					<p class="pt-5">
						{document[0].name.slice(0, 14)}
						{#if document[0].name.length > 14}
							...
						{/if}
					</p>
				{:else}
					<p class="pt-5">Seleccionar</p>
				{/if}
			</label>
		</div>
		<div class="flex-1 flex-col flex justify-center items-center">
			<IcRoundSaveAlt class="h-16 w-16" />
			<p class="pt-5 truncate">Generar</p>
		</div>
	</div>
</div>
