<script setup lang="ts">
import LandingAgents from './LandingAgents.vue';
import LandingFaq from './LandingFaq.vue';
import LandingFeatures from './LandingFeatures.vue';
import LandingFinale from './LandingFinale.vue';
import LandingHero from './LandingHero.vue';
import LandingLiveDemo from './LandingLiveDemo.vue';
import LandingQuickstart from './LandingQuickstart.vue';
</script>

<template>
	<div class="pv-landing">
		<LandingHero />
		<LandingLiveDemo />
		<LandingFeatures />
		<LandingAgents />
		<LandingQuickstart />
		<LandingFaq />
		<LandingFinale />
	</div>
</template>
