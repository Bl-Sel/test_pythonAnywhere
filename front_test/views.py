from django.shortcuts import render

# Create your views here.
def home_view(request):
    

    if request.method == "POST":
        from functions.run_macros import run_test
        ajout = run_test()
    
    return render(request, "front_test/home.html")