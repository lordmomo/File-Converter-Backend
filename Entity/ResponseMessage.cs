namespace FileConversion.Entity
{
    public class ResponseMessage
    {
        public Boolean Success { get; set; }
        public string Message {  get; set; }

        public List<string> Data { get; set; }
    }
}
