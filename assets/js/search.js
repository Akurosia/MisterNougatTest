// =============================================================================

    // APP JS

    // Authored by:
        // - Josh Beveridge

// =============================================================================

(function($) {

    $(document).ready(function() {

        var $root = $('html, body');

        // Disabled Button Clicks ==============================================
        $('.disabled').on('click', function(e){
            e.preventDefault();
            e.stopPropagation();
            return false;
        });

        // Input Focus =========================================================
        if($(".search").length) {
            $("#guideFilter").focus();
        }

        // FFXIV Guide Filters =================================================
        $("#guideFilter").keyup(function(e) {
            var input = this.value.toLowerCase().trim()
            var terms = input.split(" ");

            $("#basicAccordion").each(function(e) {
                console.log($(this))
                $(this).find(".accordion-item").each(function(e) {
                    var headingText  =  $(this).data("terms").toLowerCase();
                    var show = true;
                    terms.forEach(function(term){
                        if ( !(headingText.indexOf(term.trim()) >= 0) ) {
                            show = false;
                        }
                    });

                    if (show) {
                        $(this).show();
                        $(this).addClass("show");
                    }
                    else {
                        $(this).hide();
                        $(this).removeClass("show");
                    }

                });
            });

            // Prevents the user from submitting the form with the enter key.
            $('#guideFilter').on('keyup keypress', function(e) {
                var keyCode = e.keyCode || e.which;
                if (keyCode === 13) {

                    if ($(".site-grid__sidebar").hasClass("active")) {
                        $(".sidebar__trigger").removeClass("active");
                        $(".site-grid__sidebar-overlay").removeClass("active");
                        $(".site-grid__sidebar").removeClass("active");
                    }

                    var top = $('body').position().top;
                    $('html,body').scrollTop(top);

                    e.preventDefault();

                    document.activeElement.blur();

                    return false;

                }
            });

            e.preventDefault();

        });

        // FFXIV Guide Menu ====================================================
        $('.guide-metadata__menu-link').on("click", function(e) {

            e.preventDefault();

            var dataObject = $('#'+$(this).data('id'));
            var container = $('html,body');

            $("[class*='guide__accordion-trigger']").removeClass("active");
            $("[class*='guide__accordion-content']").removeClass("active");

            $(dataObject).parents("[class*='guide__accordion-content']").prev("[class*='guide__accordion-trigger']").addClass("active");
            $(dataObject).parents("[class*='guide__accordion-content']").addClass("active");
            $(dataObject).addClass("active");
            $(dataObject).next("[class*='guide__accordion-content']").addClass("active");

            var offSet = dataObject.offset().top;

            container.animate({
                scrollTop : offSet
            }, 'slow');

        });

        // FFXIV Guide Menu Temp ===============================================
        $('.guide-metadata__menu-link--new').on("click", function(e) {

            e.preventDefault();

            var dataTrigger = $(this).attr("data-menu-trigger");
            var dataObject = $("[data-menu-target='"+dataTrigger+"']");
            var container = $('html,body');

            $("[class*='guide__accordion-trigger']").removeClass("active");
            $("[class*='guide__accordion-content']").removeClass("active");

            $(dataObject).parents("[class*='guide__accordion-content']").prev("[class*='guide__accordion-trigger']").addClass("active");
            $(dataObject).parents("[class*='guide__accordion-content']").addClass("active");
            $(dataObject).addClass("active");
            $(dataObject).next("[class*='guide__accordion-content']").addClass("active");

            var offSet = dataObject.offset().top;

            container.animate({
                scrollTop : offSet
            }, 'slow');

        });


        // FFXIV Guide Accordions ==============================================
        $("[class*='guide__accordion-trigger']").on("click", function(e) {
            $(this).toggleClass("active");
            $(this).next("[class*='guide__accordion-content']").toggleClass("active");
            e.preventDefault();
        });

        // Guide Image Zoom ====================================================
        $(".guide__attack-image-item").on("click", function(e) {
            if($(this).hasClass("active")) {
                $(this).removeClass("active")
            }
            else {
                $(this).addClass("active")
            }
        });

    });

})(jQuery);
