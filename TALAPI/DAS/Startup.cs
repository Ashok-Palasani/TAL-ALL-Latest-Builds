using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DAS.DBModels;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.HttpsPolicy;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Swashbuckle.AspNetCore.Swagger;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.IdentityModel.Tokens;
using DAS.DAL.Helpers;
using DAS.Interface;
using DAS.DAL;


namespace DAS
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; }

        // This method gets called by the runtime. Use this method to add services to the container.
        public void ConfigureServices(IServiceCollection services)
        {
            services.Configure<CookiePolicyOptions>(options =>
            {
                // This lambda determines whether user consent for non-essential cookies is needed for a given request.
                options.CheckConsentNeeded = context => true;
                options.MinimumSameSitePolicy = SameSiteMode.None;
            });


            services.AddMvc().SetCompatibilityVersion(CompatibilityVersion.Version_2_2);
            services.AddSwaggerGen(c =>
            {
                c.SwaggerDoc("v1", new Info { Title = "DAS", Version = "v7" });
            });

            var connection = @"Server=TCP:10.80.20.25,1433;Database=i_facility_tal;user id=sa;password=password@123;";
            //var connection = @"Server=TCP:13.233.129.21,8090;Database=i_facility_tsal_New;user id=sa;password=srks4$;";
            //var connection = @"Server=DESKTOP-72HGDFG\SQLDEV17013; Database=i_facility_tal;user id=sa;password=srks4$;";
            //var connection = @"Server=TCP:10.20.10.65,1433; Database=i_facility_tsal;user id=sa;password=srks4$tsal;";
            //var connection = @"Server=TCP:SRKSDEV007\SQLEXPRESS,1236; Database=i_facility_tsal_ForTesting;user id=sa;password=srks4$;";
            services.AddDbContext<i_facility_talContext>(options => options.UseSqlServer(connection));

            services.AddCors(options =>
            {
                options.AddPolicy("AllowAllOrigins",
                    builder =>
                    {
                        builder.AllowAnyOrigin();
                        builder.AllowAnyHeader();
                        builder.AllowAnyMethod();
                    });
            });

            // configure strongly typed settings objects
            var appSettingsSection = Configuration.GetSection("AppSettings");
            services.Configure<AppSettings>(appSettingsSection);

            //// configure jwt authentication
            //var appSettings = appSettingsSection.Get<AppSettings>();
            //var key = Encoding.ASCII.GetBytes(appSettings.Secret);
            //var commonEmailKey = (appSettings.CommonEmail);
            //var documentEmail = (appSettings.DocumentEmail);
            //var resetLinkURL = (appSettings.ResetLinkURL);

            //services.AddAuthentication(x =>
            //{
            //    x.DefaultAuthenticateScheme = JwtBearerDefaults.AuthenticationScheme;
            //    x.DefaultChallengeScheme = JwtBearerDefaults.AuthenticationScheme;
            //})
            //.AddJwtBearer(x =>
            //{
            //    x.RequireHttpsMetadata = false;
            //    x.SaveToken = true;
            //    x.TokenValidationParameters = new TokenValidationParameters
            //    {
            //        ValidateIssuerSigningKey = true,
            //        IssuerSigningKey = new SymmetricSecurityKey(key),
            //        ValidateIssuer = false,
            //        ValidateAudience = false
            //    };
            //});

            #region configure DI for application services
            services.AddTransient<IPlantShopCellData, DALPlantShopCellData>();
            services.AddTransient<IHMIDetails, DALHMIDetails>();
            services.AddTransient<INoCodeInterface, DALLoss>();
            services.AddTransient<IActivity, DALActivity>();
            services.AddTransient<IBatchProcess, DALBatch>();
            services.AddTransient<IProcess, DALProcess>();
            services.AddTransient<IEmployee, DALEmployee>();
            services.AddTransient<IReport, DALReport>();
            services.AddTransient<IManualWorkCenter, ManualWorkCenterDAL>();
            services.AddTransient<IPreactorSchedule, DALPreactor>();
            //services.AddTransient<IOpCancel, DALOPCancel>();
            services.AddTransient<IHMIWrongQty, DALHMIWrongQty>();
            services.AddTransient<INoLogin, DALNoLogin>();
            services.AddTransient<ICriticalMachineMaster, CriticalMachineMasterDAL>();
            services.AddTransient<ISplitDuration, SpliDurationDAL>();
            services.AddTransient<IAndonBreakDown, DALAndonBreakDown>();
            services.AddTransient<ITcfApprovedMaster, TcfApprovedMasterDAL>();
            //services.AddTransient<IOee, OeeDAL>();

            #endregion
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IHostingEnvironment env, ILoggerFactory loggerFactory)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }
            else
            {
                app.UseExceptionHandler("/Home/Error");
                // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
                app.UseHsts();
            }

            app.UseHttpsRedirection();
            app.UseStaticFiles();
            app.UseCookiePolicy();

            // Shows UseCors with CorsPolicyBuilder.
            app.UseCors("AllowAllOrigins");

            // Enable middleware to serve generated Swagger as a JSON endpoint.
            app.UseSwagger();

            // Enable middleware to serve swagger-ui (HTML, JS, CSS, etc.), 
            // specifying the Swagger JSON endpoint.
            app.UseSwaggerUI(c =>
            {
                c.SwaggerEndpoint("/swagger/v1/swagger.json", "My API V1");
            });

            app.UseMvc(routes =>
            {
                routes.MapRoute(
                    name: "default",
                    template: "{controller=Home}/{action=Index}/{id?}");
            });

            loggerFactory.AddLog4Net();
        }
    }
}