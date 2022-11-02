using Ninject;

namespace MLineBlockAttributeEditor.Core.IoCContainer
{
    /// <summary>
    /// The IoC Container for our applications
    /// </summary>
    public static class IoC
    {
        #region PublicProperties

        /// <summary>
        /// The kernal for our IoC container
        /// </summary>
        public static IKernel Kernal { get; private set; } = new StandardKernel();
        public static bool issetup = false;
        #endregion

        #region Construction
        /// <summary>
        /// Sets up the IoC, binds all information required and is ready for use
        /// Note: Must be called as soon as the application starts up to ensure
        ///       all services can be found.
        /// </summary>
        public static void Setup()
        {
            // Bind all required viewmodels
            BindMLBEDViewModels();
        }

        /// <summary>
        /// Binds all singleton viewmodels for the Layer Controller application.
        /// </summary>
        private static void BindMLBEDViewModels()
        {
            if (issetup == false)
                // Bind to a single instance of the application view model
                Kernal.Bind<MLineBlockAttributeEditor.Core.ViewModels.MLBEDApplicationViewModel>().ToConstant(new MLineBlockAttributeEditor.Core.ViewModels.MLBEDApplicationViewModel());
            issetup = true;
        }

        /// <summary>
        /// Unbinds the bound viewmodels so the program can be launched again in the same instance of ACAD without error.
        /// </summary>
        public static void ReleaseLayerController()
        {
            // Unbind from the application view model
            //Kernal.Unbind<LayerController.LayerControllerApplicationViewModel>();
            //issetup = false;
        }
        #endregion

        #region Helper Methods
        /// <summary>
        /// gets a service from the IoC of the specified type
        /// </summary>
        /// <typeparam name="T">The type to get</typeparam>
        /// <returns></returns>
        public static T Get<T>()
        {
            return Kernal.Get<T>();
        }
        #endregion
    }
}
