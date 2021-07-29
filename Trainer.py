from chatterbot import ChatBot, filters
from chatterbot.trainers import ListTrainer
from chatterbot.trainers import ChatterBotCorpusTrainer
from chatterbot.storage import StorageAdapter
from chatterbot.conversation import Statement

import logging

# imported the logging to see Bots info
logging.basicConfig(level=logging.INFO)

# Create a new chat bot named Charlie
chatbot = ChatBot('Charlie',
                  # bot will try not to repeat responses to the user
                  # filters=[filters.get_recent_repeated_responses],
                  # SQL storage_adapters
                  storage_adapter='chatterbot.storage.SQLStorageAdapter',
                  # database
                  database_uri='sqlite:///database.sqlite3',
                  # logic adapters
                  # best_match
                  # default_response
                  logic_adapters=[
                      {
                          # using best match
                          'import_path': 'chatterbot.logic.BestMatch',
                          # default response ( low confidence )
                          'default_response': 'I am sorry, but I do not understand.',
                          # if data is 90% similar they will be clustered together
                          'maximum_similarity_threshold': 0.90
                      }
                  ],
                  # modify users input
                  # removes any extra white space
                  # cleans html characters
                  preprocessors=[
                      # removes any extra whitespaces
                      'chatterbot.preprocessors.clean_whitespace',
                      # Convert escaped html characters into unescaped html characters. For example: “&lt;b&gt;”
                      # becomes “<b>”.
                      'chatterbot.preprocessors.unescape_html'
                  ],
                  #  A filter that eliminates possibly repetitive responses to prevent
                  #     a chat bot from repeating statements that it has recently said.

                  )
corpus = ChatterBotCorpusTrainer(chatbot)

corpus.train(
   # 'chatterbot.corpus.english.issues'

)

# functions gets feed back to save user inputs and responses
def get_feedback():
    text = input()

    if 'yes' in text.lower():
        return True
    elif 'no' in text.lower():
        return False
    else:
        print('Please type either "Yes" or "No"')
        return get_feedback()


# Get a response to the input text 'I would like to book a flight.'
still_running = True
while still_running:
    input_statement = Statement(text=input())
    response = chatbot.get_response(input_statement)
    print('\n Is "{}" a  response to "{}"? \n'.format(response.text, input_statement.text))
    if get_feedback() is False:
        print('please input the correct one')
        correct_response = Statement(text=input())
        # chatbot learns the response to the users input
        chatbot.learn_response(correct_response, input_statement)
        chatbot.learn_response(correct_response, input_statement)
        # chatbot updates bot inputs from the user
        chatbot.storage.update(correct_response)

        print('Responses added to bot!')
