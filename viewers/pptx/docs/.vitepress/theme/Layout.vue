<script setup lang="ts">
import DefaultTheme from 'vitepress/theme';

import PageMotion from './PageMotion.vue';

const { Layout } = DefaultTheme;
</script>

<template>
	<Layout>
		<template #layout-top><PageMotion /></template>
	</Layout>
</template>
