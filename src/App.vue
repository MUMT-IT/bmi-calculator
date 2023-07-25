<template>
  <v-app>
    <v-main>
      <v-parallax height="100vh" src="https://cdn.vuetifyjs.com/images/backgrounds/vbanner.jpg"
        :style="{ position: 'static' }">
          <router-view />
      </v-parallax>
    </v-main>
  </v-app>
</template>
<script>

export default {
  name: 'App',

  data: () => ({
    //
  }),
};
</script>