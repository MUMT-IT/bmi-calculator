<template>
  <div>
    <v-form ref="form" lazy-validation>
      <v-card
        class="mx-auto my-12"
        max-width="370"
        height="425
  "
      >
        <v-toolbar color="indigo" dark>
          <v-toolbar-title>BMI Calculator</v-toolbar-title>
          <v-spacer></v-spacer>
        </v-toolbar>

        <v-card-text>
          <!-- Height -->
          <v-row class="mt-4">
            <v-col>
              <div align="center" class="mt-2"><b>Height</b></div>
            </v-col>
            <v-col cols="7">
              <div>
                <v-text-field
                  v-model="dataShape.height"
                  :rules="heightRules"
                  type="number"
                  autofocus
                  required
                  outlined
                  dense
                ></v-text-field>
              </div>
            </v-col>
            <v-col>
              <div class="mt-2">
                <b>cm</b>
              </div>
            </v-col>
          </v-row>

          <!-- Weight -->
          <v-row>
            <v-col>
              <div align="center" class="mt-2"><b>Weight</b></div>
            </v-col>
            <v-col cols="7">
              <div>
                <v-text-field
                  v-model="dataShape.weight"
                  :rules="weightRules"
                  type="number"
                  required
                  outlined
                  dense
                ></v-text-field>
              </div>
            </v-col>
            <v-col>
              <div class="mt-2">
                <b>kg</b>
              </div>
            </v-col>
          </v-row>

          <!-- Show BMI -->
          <v-row>
            <v-col align="center">
              <div class="fontSize20">
                <strong>Your BMI : {{ dataShape.bmi }}</strong>
                <strong :class="colorResult"> {{ result }} </strong>
              </div>
            </v-col>
          </v-row>
        </v-card-text>

        <br /><br />
        <v-divider></v-divider>

        <!-- btn  -->
        <v-card-actions>
          <v-spacer></v-spacer>
          <v-btn color="red" text @click="reset"> Clear </v-btn>
          <v-btn color="success" text @click="validate" :loading="loadingBtn">
            Calculate
          </v-btn>
        </v-card-actions>
      </v-card>

      <!-- snackbar result pust data to db -->
      <v-snackbar v-model="snackBar.showSnackBar" :timeout="3000" top>
        <div class="text-center">
          {{ snackBar.titleSnackBar }}
          <v-icon large class="ml-1" :color="snackBar.colorSnackBar">{{
            snackBar.iconSnackBar
          }}</v-icon>
        </div>
      </v-snackbar>
    </v-form>

    <!-- History -->
    <v-card width="400" class="mx-auto my-12">
      <v-toolbar color="indigo" dark>
        <v-toolbar-title>History</v-toolbar-title>
        <v-spacer></v-spacer>
      </v-toolbar>
      <v-data-table
        :headers="headers"
        :items="historyShape"
        :items-per-page="10"
        class="elevation-1"
      >
        <template v-slot:item="{ item }"
          ><tr align="center">
            <td>
              {{ item.height }}
            </td>
            <td>
              {{ item.weight }}
            </td>
            <td>
              {{ item.bmi }}
            </td>
          </tr></template
        >
      </v-data-table>
    </v-card>
  </div>
</template>

<script>
import { api } from "../helpers/Helpers";
export default {
  name: "bmi-calculator",
  data: () => ({
    dataShape: {
      height: null,
      weight: null,
      bmi: null,
    },
    snackBar: {
      showSnackBar: false,
      titleSnackBar: "",
      colorSnackBar: "",
      iconSnackBar: "",
    },
    loadingBtn: false,

    historyShape: [],
    result: "",
    colorResult: "",
    heightRules: [
      (v) => !!v || "Height is required",
      (v) => v >= 1 || "Height is more than 1",
    ],
    weightRules: [
      (v) => !!v || "Weight is required",
      (v) => v >= 1 || "Weight is more than 1",
    ],

    headers: [
      {
        text: "Height",
        align: "center",
        value: "height",
      },
      { text: "Weight", align: "center", value: "weight" },
      { text: "Bmi", align: "center", value: "bmi" },
    ],
  }),

  async mounted() {
    this.historyShape = await api.gettasks();
  },

  methods: {
    alertShow(show, title, color, icon) {
      this.snackBar = {
        showSnackBar: show,
        titleSnackBar: title,
        colorSnackBar: color,
        iconSnackBar: icon,
      };
    },

    reset() {
      this.dataShape.bmi = null;
      this.result = null;
      this.$refs.form.reset();
    },

    async validate() {
      this.$refs.form.validate();
      if (
        !!this.dataShape.height &&
        !!this.dataShape.weight &&
        this.dataShape.height >= 1 &&
        this.dataShape.weight >= 1
      ) {
        this.dataShape.bmi = parseFloat(
          (
            this.dataShape.weight /
            ((this.dataShape.height * this.dataShape.height) / 10000)
          ).toFixed(1)
        );
        this.loadingBtn = true;
        if (this.dataShape.bmi < 18.5) {
          this.colorResult = "orange--text";
          this.result = "(Underweight)";
        } else if (this.dataShape.bmi >= 18.5 && this.dataShape.bmi <= 24.9) {
          this.colorResult = "green--text";
          this.result = "(Normal Weight)";
        } else if (this.dataShape.bmi >= 25 && this.dataShape.bmi <= 29.9) {
          this.colorResult = "orange--text";
          this.result = "(Overweight)";
        } else {
          this.colorResult = "red--text";
          this.result = "(Obesity)";
        }

        // add data to DB มี backend เปิด comment
        const result = await api.createtask(this.dataShape);

        setTimeout(async () => {
          if (result._id) {
            this.alertShow(true, "Success", "success", "mdi-shield-check");
            this.historyShape = await api.gettasks();
            this.loadingBtn = false;
          } else {
            this.alertShow(true, "Fail", "red", "mdi-shield-alert");
            this.loadingBtn = false;
          }
        }, 1500);
      }
    },
  },
};
</script>

<style>
.fontSize20 {
  font-size: 20px;
}
</style>
