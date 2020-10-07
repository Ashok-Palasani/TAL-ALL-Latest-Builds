using DAS.EntityModels;
using System;
using System.Collections.Generic;
using System.Text;

namespace DAS.Interface
{
   public interface IEmployee
    {
        CommonResponse GetEmployee();
        EntityModel CreateEmployee(EntityEmployee data);
        CommonResponse EditEmployee(int id);
        //CommonResponse UpdateEmployee(EntityEmployee data);
        CommonResponse DeleteEmployee(int id);
        CommonResponse GetRole();
    }
}
