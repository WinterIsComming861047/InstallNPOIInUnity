using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using prjCGSSim;
public class init : MonoBehaviour
{
    // Start is called before the first frame update
    void Start()
    {
        List<Dictionary<string, object>> ExcelData = ExcelManager.readExcel();
        foreach (Dictionary<string, object> dataAll in ExcelData)
        {
            string tempArea="";
            foreach ( KeyValuePair< string, object> dataSingle in dataAll)
            {
                if (tempArea != ""){tempArea =  tempArea+ " , "; }
                tempArea = tempArea+ $"{dataSingle.Key} : {dataSingle.Value}";

            }


            Debug.Log(tempArea);
        }
    }

    // Update is called once per frame
    void Update()
    {
        
    }
}
