<template>
  <v-container>
    <v-row class="text-center">
      <v-col cols="12">
        <v-img
          src='https://mangeet02.github.io/delete-mail-dialog-outlook-add-in/assets/icon.png'
          class="my-3"
          contain
          height="200"
        />
      </v-col>

      <v-col class="mb-4">
        <h1 class="display-1 font-weight-bold mb-3">
          Move this mail in the junk folder
        </h1>

        <p class="subheading font-weight-regular">
          
        </p>
      </v-col>

      <v-col id="button" class="mb-5" cols="12">
        <v-btn 
          color="red"  
          class="font-weight-bold"
        >
          MOVE
        </v-btn>
      </v-col>
    </v-row>
  </v-container>
</template>

<script>
/* eslint-disable */
export default {
  name: 'HelloWorld',
  mounted() {
    Office.onReady(() => {
      document.querySelector("#button").addEventListener("click", () => this.sendMessageToParent("move"));
    });
  },
  methods: {
    sendMessageToParent(messageToSend) {
      console.log(`Sending message "${messageToSend}" from prompt dialog via messageParent method`);
      Office.context.ui.messageParent(messageToSend, { targetOrigin: "*" });
    },
  },
};
</script>
