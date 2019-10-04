using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ProjectService.Models
{
    public class Result
    {
        public object data { get; set; }
        public bool succeeded { get; set; }
        public int code { get; set; }
        public string message { get; set; }

        public Result()
        {
            this.succeeded = false;
            this.code = -1;
            this.message = "";
            this.data = null;
        }

        public Result(Exception ex)
        {
            this.data = "";
            this.message = "Error >>> " + ex.Message;
            this.code = 500;
            this.succeeded = false;
        }

        public Result ReturnDefault()
        {
            return this;
        }
    }
}