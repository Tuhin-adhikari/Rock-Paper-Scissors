import time
import win32com.client as w
import random

t=int(time.strftime("%H"))
s=w.Dispatch("SAPI.SpVoice")
s.speak("Hello user,Please enter your name");
name=input("Enter your name:\n")

if t>=0 and t<12:
    s.speak(f"Good morning {name}")

elif t>=12 and t<16:
    s.speak(f"Good afternoon {name}")

elif t>=16 and t<19:
    s.speak(f"Good evening {name}")

else:
    s.speak(f"Good past evening {name}")

s.speak("Time to play rock paper scissor. Select 1 for rock , 2 for paper and 3 for scissor ")
print("Select \n 1-> Rock \n 2-> Paper \n 3->Scissor")

comp_score = 0
my_score = 0

def play():
    user=int(input("Enter your choice:\n"))
    comp=random.randint(1,3)

    s.speak("Your choice was Rock") if user==1 else s.speak("your choice was paper") if user==2 else s.speak("Your choice was scissor") if user==3 else s.speak("Wrong input")
    s.speak("My choice was Rock") if comp==1 else s.speak("My choice was paper") if comp==2 else s.speak("My choice was scissor")

    if user==1 or user==2 or user==3:
        if user==comp:
            return 0
        
        elif user==1 and comp==2:
            return -1
        
        elif user==1 and comp==3:
            return 1
        
        elif user==2 and comp==1:
            return 1
        
        elif user==2 and comp==3:
            return -1
        
        elif user==3 and comp==1:
            return -1
        
        elif user==3 and comp==2:
            return 1
        
    else:
        s.speak("Sorry that was a wrong input!")
        exit()

while True:
    score=play()
    if score==1:
        s.speak("You win.")
        my_score+=1
    elif score==-1:
        s.speak("I win.")
        comp_score+=1
    else:
        s.speak("It was a draw, let's play again!")
        my_score+=0.5
        comp_score+=0.5

    s.speak("Enter yes to play again")
    command=input("Enter yes to play again or no to exit:\n")
    if command != "yes":
        print(f"Your score : {my_score}")
        print(f"My score : {comp_score}")
        s.speak(f"The final scores are as follows\n you scored {my_score} points and i scored {comp_score} points")
        if my_score > comp_score:
            s.speak("You are the ultimate winner")
        elif comp_score > my_score:
            s.speak("I am the ultimate winner")
        else:
            s.speak("It's an ultimate draw")
        break