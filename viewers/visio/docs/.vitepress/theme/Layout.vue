<script setup lang="ts">
import DefaultTheme from 'vitepress/theme';

import DocsSidebarTop from './DocsSidebarTop.vue';
import PageMotion from './PageMotion.vue';

const { Layout } = DefaultTheme;
</script>

<template>
	<Layout>
		<template #layout-top><PageMotion /></template>
		<template #sidebar-nav-before><DocsSidebarTop letter="V" /></template>
	</Layout>
</template>
