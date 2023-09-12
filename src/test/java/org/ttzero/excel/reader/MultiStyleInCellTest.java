/*
 * Copyright (c) 2017-2019, guanquan.wang@yandex.com All Rights Reserved.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package org.ttzero.excel.reader;

import org.junit.Test;

import java.io.IOException;
import java.util.Arrays;
import java.util.Objects;

import static org.ttzero.excel.reader.ExcelReaderTest.testResourceRoot;

/**
 * @author guanquan.wang at 2019-06-12 17:26
 */
public class MultiStyleInCellTest {
    @Test public void testMulti() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("multi-style In cell.xlsx"))) {
            String[] array = reader.sheet(0).rows().map(row -> row.getString(2)).filter(Objects::nonNull).toArray(String[]::new);
            String[] expectArray = {"abred中文","IM8zc","B6tw","JWGv25V","fTY56z","nxWnqE","RyyJ8o","UgkTgnx","JDrYdU","Bl7Lh","w2vc9","4xbrwu","5A8RN","LUYhEG","y1Ee5Sl","fyM1","p6Dn","tMMWp","P1coCD","Ej2vXbZ","aUkcla","z2aLkN8","ljjaD8","r6T1hEP","qu5iO","TJt8C","sJlIYKN","Lt9NMp","D7bVWFk","1BwJHPP","S7Pf8hG","bHXsj49","EqXOeS","exyeBe","FtKjhf","KVs8LWt","uEkSmmj","1DxEoJ","tP3d","Nb2fu","nPJ6","2CrTk17","JWbwVD","aaX7p","NkFL","9j0xsxt","6rcGpax","h0blSId","gCPR","JO1rz","m5ZEpB3","WPpB77c","iEct","Tfuq","cHv9n3Y","TX9dhf","1Ueoj7","1EY9V","0ruA","XCXhJRp","uD36","Iat3fJE","1YsnAY","yrN9e","sSIM7G","gDXE","uamheu","ZXMN","8MymZBX","ic9aLbJ","e1QuI","bEQYyk","3DKbi0k","NRjW2","vz9HLE","aBYIMi","7Xe2lQz","d95SDS","jUfZ","ODrB2"};
            assert Arrays.equals(expectArray, array);
        }
    }
}
