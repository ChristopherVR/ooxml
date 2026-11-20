<script setup lang="ts">
import type { SignatureStatusView } from 'ooxml-ui/pptx';
import { useI18n } from 'vue-i18n';

import {
	SIGNATURE_BADGE_TONE,
	SIGNATURE_HEADER_TONE,
	SIGNATURE_ICONS,
	SIGNATURE_ROW_TONE,
} from './signature-look';

const { t } = useI18n();

/**
 * SignaturesPanel - read-only digital-signature status panel.
 *
 * Lists each signature in the package (signer, its verification state, issuer,
 * signing time) under the deck's overall state. Every word and colour comes
 * from the shared `signatureStatusView`, so a deck reads "Signed" only when its
 * signatures verify. Purely presentational: it inspects, it never signs or strips.
 */
const props = defineProps<{
	view: SignatureStatusView;
}>();
</script>

<template>
	<section
		class="pptx-vue-signatures overflow-hidden rounded-lg border border-border bg-popover text-[13px] text-foreground"
		:aria-label="t('pptx.digitalSignatures.ariaLabel')"
		:data-pptx-signature-state="props.view.state"
	>
		<header
			class="pptx-vue-signatures__header flex items-center gap-2 border-b border-border px-3 py-2.5 font-semibold"
			:class="[
				`pptx-vue-signatures__header--${props.view.state}`,
				SIGNATURE_HEADER_TONE[props.view.tone],
			]"
		>
			<component
				:is="SIGNATURE_ICONS[props.view.icon]"
				class="pptx-vue-signatures__icon h-4 w-4 flex-none"
				aria-hidden="true"
			/>
			<span class="pptx-vue-signatures__title flex-1">{{
				props.view.present ? props.view.label : t('pptx.digitalSignatures.notSigned')
			}}</span>
			<span
				v-if="props.view.present"
				class="pptx-vue-signatures__count text-xs font-normal opacity-80"
			>
				{{ t('pptx.digitalSignatures.signatureCount', { count: props.view.count }) }}
			</span>
		</header>

		<p
			v-if="!props.view.present"
			class="pptx-vue-signatures__empty m-0 px-3 py-3.5 text-muted-foreground"
		>
			{{ t('pptx.digitalSignatures.noSignatures') }}
		</p>

		<template v-else>
			<p class="pptx-vue-signatures__description m-0 px-3 pt-2.5 text-xs text-muted-foreground">
				{{ props.view.description }}
			</p>
			<ul class="pptx-vue-signatures__list m-0 list-none p-0">
				<li
					v-for="row in props.view.rows"
					:key="row.path"
					class="pptx-vue-signatures__item border-b border-l-[3px] border-b-border/60 px-3 py-2.5 last:border-b-0"
					:class="[`pptx-vue-signatures__item--${row.tone}`, SIGNATURE_ROW_TONE[row.tone]]"
					:data-pptx-signature-row-state="row.state"
				>
					<div class="pptx-vue-signatures__item-main flex items-center justify-between gap-2">
						<span
							class="pptx-vue-signatures__signer font-semibold break-words"
							:title="row.subject"
							>{{ row.signer }}</span
						>
						<span
							class="pptx-vue-signatures__badge inline-flex flex-none items-center gap-1 whitespace-nowrap rounded-full border px-2 py-0.5 text-[11px] font-semibold"
							:class="[`pptx-vue-signatures__badge--${row.tone}`, SIGNATURE_BADGE_TONE[row.tone]]"
						>
							<component :is="SIGNATURE_ICONS[row.icon]" class="h-3 w-3" aria-hidden="true" />
							{{ row.label }}
						</span>
					</div>

					<dl
						v-if="row.issuer || row.signedAt"
						class="pptx-vue-signatures__meta m-0 mt-1.5 grid grid-cols-[auto_1fr] gap-x-2.5 gap-y-0.5 text-xs text-muted-foreground [&_dd]:m-0 [&_dd]:break-words [&_dt]:font-medium [&_dt]:text-muted-foreground"
					>
						<template v-if="row.issuer">
							<dt>{{ t('pptx.digitalSignatures.issuer') }}</dt>
							<dd>{{ row.issuer }}</dd>
						</template>
						<template v-if="row.signedAt">
							<dt>{{ props.view.signedAtCaption }}</dt>
							<dd>{{ row.signedAt }}</dd>
						</template>
					</dl>
				</li>
			</ul>
			<p
				class="pptx-vue-signatures__removed-on-save m-0 border-t border-border px-3 py-2 text-[11px] text-muted-foreground"
			>
				{{ props.view.removedOnSave }}
			</p>
		</template>
	</section>
</template>
