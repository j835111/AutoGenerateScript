namespace WindowsFormsApplication5
{
    //存輸入數據
    public class InputData
    {
        public int number;                           //輸入編號
        public static int max = 1000;
        string instrument;                           //儀器名稱
        string function;                             //函式名稱
        string content;                              //內容
        string[] parameter = new string[max];        //參數
        string[] pin_name = new string[max];         //腳位
        string[] relay = new string[max];            //連接relay
        string[] com1 = new string[max];             //com1
        string[] com2 = new string[max];             //com2
        public string delaySec;                      //delay秒數

        public string GetInstrument { get => instrument; }
        public string GetFunction { get => function; }
        public string GetContent { get => content; }
        public string[] GetParameter { get => parameter; }
        public string[] GetPin_name { get => pin_name; }
        public string[] GetRelay { get => relay; }
        public string[] GetCom1 { get => com1; }
        public string[] GetCom2 { get => com2; }
        public string GetNumber { get => number.ToString(); }
        public string GetDelaySec()
        {
            return delaySec;
        }

        public InputData(int number)
        {
            this.number = number;
        }
        public void SetData(string instrument, string function, string content, string parameter, string delaySec)
        {
            this.delaySec = delaySec;
            this.instrument = instrument;
            this.function = function;
            this.content = content;
            int i = 0;
            foreach (string str in parameter.Split(','))
            {
                this.parameter[i] = str;
                i++;
            }
        }
        public void SetData(int i, string pin_name, string relay, string com1, string com2)
        {
            this.pin_name[i] = pin_name;
            this.relay[i] = relay;
            this.com1[i] = com1;
            this.com2[i] = com2;
        }  
    }
}
