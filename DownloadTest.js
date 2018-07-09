/* @flow */

import React, { Component } from "react";
import { View, StyleSheet, TouchableOpacity, Alert } from "react-native";
import { Button, Text } from "native-base";
import XLSX from "xlsx";
import {
  writeFile,
  ExternalStorageDirectoryPath,
  mkdir
} from "react-native-fs";

//ReactNative FS
const DDP = ExternalStorageDirectoryPath + "/ReactNative/";
const output = str => str;

// const make_cols = refstr =>
//   Array.from({ length: XLSX.utils.decode_range(refstr).e.c + 1 }, (x, i) =>
//     XLSX.utils.encode_col(i)
//   );

export default class DownloadTest extends Component {
  constructor(props) {
    super(props);
    this.state = {
      data: [
        ["Name", "Age", "Address"],
        ["Bond", 35, "Tirunelveli"],
        ["Edison", 15, "Tirunelveli"],
        ["Vargess", 35, "Coimbatore"]
      ],
      hotelData: [
        ["HotelName", "City"],
        ["Marriot", "Coimbatore"],
        ["LeMeriden", "Coimbatore"]
      ],
      widthArr: [60, 60, 60]
      // cols: make_cols("A1:C2")
    };
    this.exportFile = this.exportFile.bind(this);
  }

  componentWillMount() {
    console.log("Component WILL MOUNT!");
    const DIR = mkdir(ExternalStorageDirectoryPath + "/ReactNative");
  }

  exportFile() {
    const ws = XLSX.utils.aoa_to_sheet(this.state.data);

    const wt = XLSX.utils.aoa_to_sheet(this.state.hotelData);
    //build new work book
    const wb = XLSX.utils.book_new();

    // ws.write_formula("B6", "=SUM(B2:B4)");
    XLSX.utils.book_append_sheet(wb, ws, "sheetjs");
    XLSX.utils.book_append_sheet(wb, wt, "sheetTwojs");

    ws["B6"].f = "B2 + B3";
    console.log(ws);

    //Write File
    const wbout = XLSX.write(wb, { type: "binary", bookType: "xlsx" });
    const file = DDP + "sheetjs.xlsx";

    writeFile(file, output(wbout), "ascii")
      .then(res => {
        Alert.alert("exportFile sucess", "Exported to" + file);
      })
      .catch(err => {
        Alert.alert("exportFile error", "error message to" + err.message);
      });
  }

  render() {
    return (
      <View style={styles.container}>
        <Text>I'm the DownloadTest component</Text>
        <Button onPress={this.exportFile}>
          <Text>export data</Text>
        </Button>
      </View>
    );
  }
}

const styles = StyleSheet.create({
  container: {
    flex: 1,
    alignItems: "center"
  }
});
