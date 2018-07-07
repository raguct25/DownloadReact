/* @flow */

import React, { Component } from "react";
import { View, Text, StyleSheet, TouchableOpacity, Alert } from "react-native";
import XLSX from "xlsx";
import {
  writeFile,
  ExternalStorageDirectoryPath,
  mkdir
} from "react-native-fs";

const DIR = mkdir(ExternalStorageDirectoryPath + "/ReactNative");
const DDP = ExternalStorageDirectoryPath + "/ReactNative/";
const output = str => str;
const make_cols = refstr =>
  Array.from({ length: XLSX.utils.decode_range(refstr).e.c + 1 }, (x, i) =>
    XLSX.utils.encode_col(i)
  );

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

  exportFile() {
    const ws = XLSX.utils.aoa_to_sheet(this.state.data);
    const wt = XLSX.utils.aoa_to_sheet(this.state.hotelData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "sheetjs");
    XLSX.utils.book_append_sheet(wb, wt, "sheetTwojs");
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
        <TouchableOpacity>
          <Text onPress={this.exportFile}> export data </Text>
        </TouchableOpacity>
      </View>
    );
  }
}

const styles = StyleSheet.create({
  container: {
    flex: 1
  }
});
