using Google.Maps;
using Google.Maps.Geocoding;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Google_Maps_Address_Qualifier
{
    class Parsers
    {

        public static List<Result> CheckAddress(string address)
        {
            List<Result> output = new List<Result>();

            var request = new GeocodingRequest();
            request.Address = address;
            request.Sensor = false;
            var response = new GeocodingService().GetResponse(request);

            if (response.Status == ServiceResponseStatus.Ok)
            {
                var tmpOut = response.Results.ToList();

                foreach (var element in tmpOut)
                {
                    output.Add(element);
                }
            }
            else if (response.Status == ServiceResponseStatus.ZeroResults)
            {
                Result NotFound = new Result();
                output.Add(NotFound);
            }
            else if (response.Status == ServiceResponseStatus.OverQueryLimit)
            {
                MessageBox.Show(response.Status.ToString(), "Error");
            }
            else
            {
                MessageBox.Show(response.Status.ToString(), "Error");
            }

            return output;
        }

    }
}
