import pandas as pd
import numpy as np
from sklearn import model_selection
from sklearn import metrics
from keras.layers import Input, Dense, LSTM, TimeDistributed, Lambda, multiply
from keras.models import Model, load_model
from keras.optimizers import Adam
from keras.preprocessing.sequence import pad_sequences
from keras import backend as K

import operator


def generate_datasets():
    """
            skill_builder_data_corrected.csv:原始数据
    """
    filename = r'D:\Program Files\RNNModelingForEdu-master\RNNModelingForEdu-master\skill_builder_data_corrected.csv'
    df = pd.read_csv(filename, encoding='utf-8', low_memory=False)
    df = df[(df['original'] == 1) & (df['attempt_count'] == 1) & ~(df['skill_name'].isnull())]
    users_list = df['user_id'].unique()
    skill_list = df['skill_name'].unique()
    skill_dict = dict(zip(skill_list, np.arange(len(skill_list), dtype='int32') + 1))  # 生成知识点字典
    response_list = []
    skill_list = []
    assistment_list = []

    counter = 0
    for user in users_list:
        sub_df = df[df['user_id'] == user]
        if len(sub_df) > 100:
            first_hundred = sub_df.iloc[0:100]
            response_df = pd.DataFrame(index=[counter], columns=['student_id'] + ['r' + str(i) for i in range(100)])
            skill_df = pd.DataFrame(index=[counter], columns=['student_id'] + ['s' + str(i) for i in range(100)])
            assistment_df = pd.DataFrame(index=[counter], columns=['student_id'] + ['a' + str(i) for i in range(100)])

            response_df.iloc[0, 0] = first_hundred.iloc[0]['user_id']
            skill_df.iloc[0, 0] = first_hundred.iloc[0]['user_id']
            assistment_df.iloc[0, 0] = first_hundred.iloc[0]['user_id']
            for i in range(100):
                response_df.iloc[0, i + 1] = first_hundred.iloc[i]['correct']
                skill_df.iloc[0, i + 1] = skill_dict[first_hundred.iloc[i]['skill_name']]
                assistment_df.iloc[0, i + 1] = first_hundred.iloc[i]['assistment_id']
            counter += 1
            response_list.append(response_df)
            skill_list.append(skill_df)
            assistment_list.append(assistment_df)

    response_df = pd.concat(response_list)  # 每个学生做的题目的掌握情况；1：正确；0：错误
    skill_df = pd.concat(skill_list)  # 每个学生做题对应的每个知识点编号
    assistment_df = pd.concat(assistment_list)  # 每个学生的做题编号

    return skill_dict, response_df, skill_df, assistment_df

def read_file(self):
    # 从文件中读取数据，返回读取出来的数据和知识点个数
    # 保存每个学生的做题信息 {学生id: [[知识点id，答题结果], [知识点id，答题结果], ...]}，用一个二元列表来表示一个学生的答题信息
    seqs_by_student = {}
    skills = []  # 统计知识点的数量，之后输入的向量长度就是两倍的知识点数量
    with open(self.fileName, 'r') as f:
        for line in f:
            fields = line.strip().split(" ")  # 一个列表，[学生id，知识点id，答题结果]
            student, skill, is_correct = int(fields[0]), int(fields[1]), int(fields[2])
            skills.append(skill)  # skill实际上是用该题所属知识点来表示的
            seqs_by_student[student] = seqs_by_student.get(student, []) + [[skill, is_correct]]  # 保存每个学生的做题信息

    return seqs_by_student, list(set(skills))


def one_hot(skill_matrix, vocab_size):
    '''
    params:
        skill_matrix: 2-D matrix (student, skills)
        vocal_size: size of the vocabulary
    returns:
        a ndarray with a shape like (student, sequence_len, vocab_size)
        here is (584 students, 99 trained-skills, 110 all skills)
    '''
    seq_len = skill_matrix.shape[1]  # columns
    result = np.zeros((skill_matrix.shape[0], seq_len, vocab_size))
    for i in range(skill_matrix.shape[0]):
        result[i, np.arange(seq_len), skill_matrix[i]] = 1.
    return result


def dkt_one_hot(skill_matrix, response_matrix, vocab_size):
    seq_len = skill_matrix.shape[1]
    skill_response_array = np.zeros((skill_matrix.shape[0], seq_len, 2 * vocab_size))
    for i in range(skill_matrix.shape[0]):
        skill_response_array[i, np.arange(seq_len), 2 * skill_matrix[i] + response_matrix[i]] = 1.
    return skill_response_array  # 584*99*222


def preprocess(skill_df, response_df, skill_num):
    skill_matrix = skill_df.iloc[:, 1:].values
    response_array = response_df.iloc[:, 1:].values
    skill_array = one_hot(skill_matrix, skill_num)
    skill_response_array = dkt_one_hot(skill_matrix, response_array, skill_num)
    return skill_array, response_array, skill_response_array


def build_skill2skill_model(input_shape, lstm_dim=32, dropout=0.0):
    input = Input(shape=input_shape, name='input_skills')
    lstm = LSTM(lstm_dim,
                return_sequences=True,
                dropout=dropout,
                name='lstm_layer')(input)
    output = TimeDistributed(Dense(input_shape[-1], activation='softmax'), name='probability')(lstm)
    model = Model(inputs=[input], outputs=[output])
    adam = Adam(lr=0.001, beta_1=0.9, beta_2=0.999, epsilon=1e-07, decay=0.0)
    model.compile(optimizer=adam, loss='categorical_crossentropy', metrics=['accuracy'])
    model.summary()
    return model


def reduce_dim(x):
    x = K.max(x, axis=-1, keepdims=True)
    return x


def build_dkt_model(input_shape, lstm_dim=32, dropout=0.5):
    input_skills = Input(shape=input_shape, name='input_skills')
    lstm = LSTM(lstm_dim,
                return_sequences=True,
                dropout=dropout,
                name='lstm_layer')(input_skills)
    dense = TimeDistributed(Dense(int(input_shape[-1] / 2), activation='sigmoid'), name='probability_for_each')(lstm)

    skill_next = Input(shape=(input_shape[0], int(input_shape[1] / 2)), name='next_skill_tested')
    merged = multiply([dense, skill_next], name='multiply')
    reduced = Lambda(reduce_dim, output_shape=(input_shape[0], 1), name='reduce_dim')(merged)

    model = Model(inputs=[input_skills, skill_next], outputs=[reduced])
    adam = Adam(lr=0.001, beta_1=0.9, beta_2=0.999, epsilon=1e-07, decay=0.0)
    model.compile(optimizer=adam, loss='binary_crossentropy', metrics=['accuracy'])
    model.summary()
    return model


# def top5_hardandeasy_skills(model, skill_array, skill_num):
#     X = skill_array[:, 0:-1]
#     y = skill_array[:, 1:]
#     X_train, X_test, y_train, y_test = model_selection.train_test_split(X, y, train_size=0.7, test_size=0.3)
#
#     model.fit(X_train,
#               y_train,
#               epochs=20,
#               batch_size=32,
#               shuffle=True,
#               validation_split=0.2)
#
#     predictions = model.predict(X_test)
#     one_hot_predictions = []
#     for i in np.arange(len(predictions)):
#         one_hot_layer = []
#         for j in np.arange(len(predictions[0])):
#             index_of_max = np.argmax(predictions[i][j])
#             one_hot_version = np.zeros(skill_num)
#             one_hot_version[index_of_max] = 1
#             one_hot_layer.append(one_hot_version)
#         one_hot_predictions.append(one_hot_layer)
#     error_rate = np.count_nonzero(y_test - one_hot_predictions) / 2 / (y_test.shape[0] * y_test.shape[1])
#     acc_rate = 1 - error_rate
#     correct = {}
#     incorrect = {}
#
#     for i in range(len(one_hot_predictions)):
#         for j in range(len(one_hot_predictions[0])):
#             prediction = one_hot_predictions[i][j]
#             actual = y_test[i][j]
#
#             comparison = prediction - actual
#             position = np.argmax(actual) + 1  # 默认按列查找最大值，返回索引（技能编号）
#             if 1 not in comparison and -1 not in comparison:  # 预测正确
#                 if position in correct.keys():
#                     correct[position] += 1
#                 else:
#                     correct[position] = 1
#             else:
#                 if position in incorrect.keys():
#                     incorrect[position] += 1
#                 else:
#                     incorrect[position] = 1
#
#     for i in np.arange(1, 112):
#         if i not in correct.keys():
#             correct[i] = 0
#         if i not in incorrect.keys():
#             incorrect[i] = 0
#
#     totals = [correct[i] + incorrect[i] for i in np.arange(1, 112)]
#     for i in range(1, len(correct) + 1):
#         if totals[i - 1] != 0:
#             correct[i] = correct[i] / totals[i - 1]
#             incorrect[i] = incorrect[i] / totals[i - 1]
#         else:
#             if correct[i] != 0 or incorrect[i] != 0:
#                 print("something is incorrect lol")
#
#     sorted_correct = sorted(correct.items(), key=operator.itemgetter(1))
#     easiest_to_identify_skills = sorted_correct[-5:]
#     sorted_incorrect = sorted(incorrect.items(), key=operator.itemgetter(1))
#     hardest_to_identify_skills = sorted_incorrect[-5:]
#
#     return easiest_to_identify_skills, hardest_to_identify_skills, acc_rate


def dkt_prediction(skill_array, skill_response_array, response_array):
    x = skill_response_array[:, 0:-1]
    skill = skill_array[:, 1:]
    response = response_array[:, 1:, np.newaxis]

    x_train, x_test, skill_train, skill_test, response_train, response_test = \
        model_selection.train_test_split(x, skill, response, train_size=0.7, test_size=0.3)
    dkt_model = load_model('dkt_model.h5')
    dkt_model.fit([x_train, skill_train],
                  response_train,
                  epochs=20,
                  batch_size=100,
                  shuffle=True,
                  validation_split=0.2)
    dkt_predictions = dkt_model.predict([x_test, skill_test])
    for i in np.arange(len(dkt_predictions)):
        for j in np.arange(len(dkt_predictions[0])):
            value = dkt_predictions[i][j][0]
            if value >= 0.5:
                dkt_predictions[i][j][0] = 1
            else:
                dkt_predictions[i][j][0] = 0

    error_rate = np.count_nonzero(response_test - dkt_predictions) / 2 / (
                response_test.shape[0] * response_test.shape[1])
    acc_rate = 1 - error_rate
    y_true = response_test.flatten()
    y_score = dkt_predictions.flatten()
    fpr, tpr, thresholds = metrics.roc_curve(y_true, y_score)
    auc = metrics.auc(fpr, tpr)
    # 所有学生知识点编号列表
    all_true_true = {}  # 原本正确预测正确
    all_true_false = {}  # 原本正确预测错误
    all_false_true = {}  # 原本错误预测正确
    all_false_false = {}  # 原本错误预测错误
    if len(dkt_predictions) == len(response_test):
        for i in range(len(dkt_predictions)):
            true_true = []  # 原本正确预测正确
            true_false = []  # 原本正确预测错误
            false_true = []  # 原本错误预测正确
            false_false = []  # 原本错误预测错误
            for j in range(len(dkt_predictions[0])):
                v_pred = dkt_predictions[i][j][0]
                v_real = response_test[i][j][0]
                if v_real:
                    true_true.append(j+1) if v_pred else true_false.append(j+1)
                else:
                    false_true.append(j+1) if v_pred else false_false.append(j+1)
            all_true_true[str(i)] = true_true
            all_true_false[str(i)] = true_false
            all_false_true[str(i)] = false_true
            all_false_false[str(i)] = false_false
    return acc_rate, auc, all_true_true, all_true_false, all_false_true, all_false_false, dkt_predictions


if __name__ == '__main__':
    skill_dict, response_df, skill_df, assistment_df = generate_datasets()

    skill_num = len(skill_dict) + 1  # including 0
    skill_array, response_array, skill_response_array = preprocess(skill_df, response_df, skill_num)

    # print('skill2skill')
    # skill2skill_model = build_skill2skill_model((99, skill_num), lstm_dim=64)
    print('dkt')
    dkt_model = build_dkt_model((99, 2 * skill_num), lstm_dim=64)
    dkt_model.save('dkt_model.h5')

    # easiest_skills, hardest_skills, acc_rate = top5_hardandeasy_skills(skill2skill_model, skill_array, skill_num)
    # print(easiest_skills, hardest_skills, acc_rate)
    acc_dkt, auc_dkt, all_true_true, all_true_false, all_false_true, all_false_false, dkt_predictions = \
        dkt_prediction(skill_array, skill_response_array, response_array)

