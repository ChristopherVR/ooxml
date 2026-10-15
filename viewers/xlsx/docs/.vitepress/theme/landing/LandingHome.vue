<script setup lang="ts">
import LandingStatus from './LandingStatus.vue';
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
		<LandingStatus />
		<LandingQuickstart />
		<LandingFaq />
		<LandingFinale />
	</div>
</template>
